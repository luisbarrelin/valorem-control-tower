'use strict';
const fs = require('node:fs');
const path = require('node:path');
const FIELDS = ['row','scheduled','arrival','operation','quantity','pallets','load','driver','plate','trailer','carrier','vehicleStatus','loadStatus','billingStatus','observation','operationEnd','statusAux','indicator','loadingStart','loadingEnd','loadingTime','palletObs'];
function decode(data) {
  if (typeof data !== 'string' || Buffer.byteLength(JSON.stringify({data})) > 64000) throw new Error('Payload ausente ou maior que o limite.');
  const input = JSON.parse(data);
  if (input.format !== 'valorem-report-v1' || input.sheet !== 'Report' || JSON.stringify(input.fields) !== JSON.stringify(FIELDS) || !Array.isArray(input.strings) || !Array.isArray(input.rows) || !input.rows.length || input.rows.length > 10000 || !Number.isFinite(Date.parse(input.sourceModifiedAt))) throw new Error('Formato do Report incompatível.');
  if (!input.strings.every(value => typeof value === 'string')) throw new Error('Dicionário de textos inválido.');
  const seen = new Set();
  const records = input.rows.map(row => {
    if (!Array.isArray(row) || row.length !== FIELDS.length || !Number.isSafeInteger(row[0]) || row[0] < 1 || seen.has(row[0])) throw new Error('Linha inválida ou duplicada.');
    seen.add(row[0]);
    const record = Object.fromEntries(FIELDS.map((field, index) => {
      let value = row[index];
      if (Array.isArray(value)) {
        if (value.length !== 1 || !Number.isSafeInteger(value[0]) || value[0] < 0 || value[0] >= input.strings.length) throw new Error('Referência de texto inválida.');
        value = input.strings[value[0]];
      } else if (value !== null && typeof value !== 'boolean' && !(typeof value === 'number' && Number.isFinite(value))) throw new Error('Valor de célula inválido.');
      return [field, value];
    }));
    if (record.scheduled == null && record.load == null) throw new Error('Registro sem agendamento e carga.');
    return record;
  });
  return { version: 1, sourceFile: String(input.sourceFile || 'Report.xlsx'), sheet: 'Report', importedAt: input.sourceModifiedAt, sourceModifiedAt: input.sourceModifiedAt, syncOrigin: 'sharepoint', records };
}
function publish(eventPath, outputPath) {
  const event = JSON.parse(fs.readFileSync(eventPath, 'utf8'));
  if (event.action !== 'sharepoint-report') throw new Error('Evento de sincronização inválido.');
  const dataset = decode(event.client_payload?.data);
  if (fs.existsSync(outputPath)) {
    const previous = JSON.parse(fs.readFileSync(outputPath, 'utf8'));
    if (Date.parse(previous.importedAt) > Date.parse(dataset.importedAt)) {
      console.log('Evento antigo ignorado; a versão mais recente foi preservada.');
      return false;
    }
    if (JSON.stringify(previous.records) === JSON.stringify(dataset.records)) {
      console.log('Report sem alterações de dados.');
      return false;
    }
  }
  fs.mkdirSync(path.dirname(outputPath), {recursive:true});
  const temporary = outputPath + '.tmp';
  fs.writeFileSync(temporary, JSON.stringify(dataset));
  fs.renameSync(temporary, outputPath);
  console.log(`Report validado: ${dataset.records.length} registros.`);
  return true;
}
if (require.main === module) publish(process.env.GITHUB_EVENT_PATH, path.resolve('data/report.json'));
module.exports = {decode, publish, FIELDS};
