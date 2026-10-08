import fs from 'node:fs/promises';
import { SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const outDir = 'outputs';
await fs.mkdir(outDir, { recursive: true });
const wb = Workbook.create();
const readme = wb.worksheets.add('Readme');
const personas = wb.worksheets.add('Personas');
const servicios = wb.worksheets.add('Servicios');
const asignaciones = wb.worksheets.add('Asignaciones');
const catalogos = wb.worksheets.add('Catalogos');
const blue = '#123B66';
const light = '#EAF2FB';
const border = '#D9E2EC';
const font = { name: 'Aptos', size: 11 };

readme.getRange('A1:F12').values = [
  ['Base de datos local — Personas y servicios PID', null, null, null, null, null],
  ['Propósito', 'Libro Excel para mantener personas, servicios PID y asignaciones de responsables.', null, null, null, null],
  ['Flujo previsto', 'La aplicación local importará este libro y exportará los cambios para mantenerlo actualizado.', null, null, null, null],
  ['Personas', 'Una fila por persona. Tipo: SEGITTUR, INECO o Externo. Área solo aplica a SEGITTUR.', null, null, null, null],
  ['Asignaciones', 'Una fila por relación persona-servicio. Una persona puede aparecer en varios servicios.', null, null, null, null],
  ['Rol', 'Principal o Corresponsable. Se recomienda un único responsable principal por servicio.', null, null, null, null],
  ['Identificadores', 'PersonaID, ServicioID y AsignacionID son claves estables para sincronizar con la aplicación.', null, null, null, null],
  ['Fecha de actualización', new Date(), null, null, null, null],
  ['Fuente de servicios', 'Catálogo PID consultado en la aplicación SEGITTUR.', null, null, null, null],
  ['Pendiente', 'Completar descripciones desde las URLs de detalle que se vayan facilitando.', null, null, null, null],
  ['Nota', 'Este libro inicial no contiene todavía personas ni asignaciones reales.', null, null, null, null],
  ['Versión', '1.0', null, null, null, null],
];
readme.mergeCells('A1:F1');
readme.getRange('A1:F1').format = { fill: blue, font: { name: 'Aptos', size: 15, bold: true, color: '#FFFFFF' } };
readme.getRange('A2:A12').format = { fill: light, font: { name: 'Aptos', size: 11, bold: true, color: blue } };
readme.getRange('A1:F12').format.font = font;
readme.getRange('B8').format.numberFormat = 'yyyy-mm-dd';

const peopleHeaders = ['PersonaID','Nombre','Apellidos','Email','TipoPersona','AreaSEGITTUR','Activo','Notas'];
personas.getRange('A1:H1').values = [peopleHeaders];
personas.getRange('A2:H2').values = [['PER-0001','','','','SEGITTUR','','Sí','']];
personas.getRange('A1:H2').format.font = font;
personas.getRange('A1:H1').format = { fill: blue, font: { name: 'Aptos', size: 11, bold: true, color: '#FFFFFF' }, wrapText: true };
personas.tables.add('A1:H2', true, 'PersonasTable');
personas.getRange('E2:E500').dataValidation = { rule: { type: 'list', values: ['SEGITTUR','INECO','Externo'] } };
personas.getRange('G2:G500').dataValidation = { rule: { type: 'list', values: ['Sí','No'] } };

const serviceNames = ['Portal Profesional','CMS Semántico','Portal Destino','Landing Page','Inteligencia de la gestión de interacción y fidelización','Reserva online de Servicios Turísticos','App Turista','Asistente Virtual Cognitivo','Gestión del Destino DTI','Participación de empresas turísticas','Inteligencia de localizaciones','Índice de madurez digital de las PYMES','SIT Destino','Encuestas en Destino','Reputación online','Análisis de percepción ciudadana sobre el Destino','Explotación de datos, previsión, simulación','Rendimiento DTI','Datos del Destino para empresas','Planificador de Rutas','Gestión de la planificación turística','Lector Validador de Códigos QR','Gestor de marketing','Linx - Espacio de Datos de Turismo'];
const serviceRows = serviceNames.map((name, i) => [`SER-${String(i+1).padStart(3,'0')}`, name, name === 'Portal Profesional' ? 'El portal centraliza los flujos de trabajo profesionales de la Plataforma Inteligente de Destino: gestión de empresas, instalaciones turísticas, servicios, eventos, ofertas y moderación de destinos.' : 'Pendiente de extracción', '', 'Activo']);
servicios.getRange(`A1:E${serviceRows.length+1}`).values = [ ['ServicioID','Servicio','Descripción','URLDetalle','Estado'], ...serviceRows ];
servicios.getRange(`A1:E${serviceRows.length+1}`).format.font = font;
servicios.getRange('A1:E1').format = { fill: blue, font: { name: 'Aptos', size: 11, bold: true, color: '#FFFFFF' }, wrapText: true };
servicios.tables.add(`A1:E${serviceRows.length+1}`, true, 'ServiciosTable');
servicios.getRange(`E2:E${serviceRows.length+1}`).dataValidation = { rule: { type: 'list', values: ['Activo','Inactivo'] } };

asignaciones.getRange('A1:H1').values = [['AsignacionID','PersonaID','ServicioID','Rol','FechaInicio','FechaFin','Activo','Notas']];
asignaciones.getRange('A2:H2').values = [['ASI-0001','','','Principal',null,null,'Sí','']];
asignaciones.getRange('A1:H2').format.font = font;
asignaciones.getRange('A1:H1').format = { fill: blue, font: { name: 'Aptos', size: 11, bold: true, color: '#FFFFFF' }, wrapText: true };
asignaciones.tables.add('A1:H2', true, 'AsignacionesTable');
asignaciones.getRange('D2:D500').dataValidation = { rule: { type: 'list', values: ['Principal','Corresponsable'] } };
asignaciones.getRange('G2:G500').dataValidation = { rule: { type: 'list', values: ['Sí','No'] } };
asignaciones.getRange('E2:F500').format.numberFormat = 'yyyy-mm-dd';

catalogos.getRange('A1:D5').values = [
  ['TipoPersona','AreaSEGITTUR','Rol','Estado'],
  ['SEGITTUR','Área por definir','Principal','Activo'],
  ['INECO','No aplica','Corresponsable','Inactivo'],
  ['Externo','','',''],
  ['','','',''],
];
catalogos.getRange('A1:D5').format.font = font;
catalogos.getRange('A1:D1').format = { fill: blue, font: { name: 'Aptos', size: 11, bold: true, color: '#FFFFFF' } };

for (const sheet of [readme, personas, servicios, asignaciones, catalogos]) {
  sheet.showGridLines = false;
  sheet.getUsedRange()?.format.autofitColumns();
  sheet.freezePanes.freezeRows(1);
}
personas.getRange('A1:H500').format.wrapText = true;
servicios.getRange('A1:E100').format.wrapText = true;
asignaciones.getRange('A1:H500').format.wrapText = true;
readme.getRange('A:A').format.columnWidth = 24;
readme.getRange('B:B').format.columnWidth = 85;
personas.getRange('B:C').format.columnWidth = 18;
personas.getRange('H:H').format.columnWidth = 32;
servicios.getRange('B:B').format.columnWidth = 38;
servicios.getRange('C:C').format.columnWidth = 72;
servicios.getRange('D:D').format.columnWidth = 48;
asignaciones.getRange('H:H').format.columnWidth = 32;

wb.recalculate();
const check = await wb.inspect({ kind: 'table', sheetId: 'Personas', range: 'A1:H4', include: 'values,formulas', tableMaxRows: 4, tableMaxCols: 8, maxChars: 3000 });
console.log(check.ndjson);
const preview = await wb.render({ sheetName: 'Personas', range: 'A1:H8', scale: 1, format: 'png' });
await fs.writeFile(`${outDir}/personas_preview.png`, new Uint8Array(await preview.arrayBuffer()));
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(`${outDir}/personas_servicios_pid.xlsx`);
