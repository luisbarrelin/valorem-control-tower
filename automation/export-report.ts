// Office Script: reads Report and returns the dashboard fields. Never changes cells.
function main(workbook: ExcelScript.Workbook, modifiedAt: string = ''): string {
  const sheet = workbook.getWorksheet('Report');
  if (!sheet) throw new Error('A aba Report não foi encontrada.');
  const range = sheet.getUsedRange(true);
  if (!range) throw new Error('A aba Report está vazia.');
  // Resolve Excel calls before array callbacks: the cloud runner cannot rewrite
  // findIndex callbacks that access the workbook into async callbacks.
  const startRow = range.getRowIndex();
  const startColumn = range.getColumnIndex();
  const values = range.getValues();
  const normalize = (value: string | number | boolean): string => String(value ?? '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase();
  const header = values.findIndex((row) => row.some((cell) => normalize(cell).includes('AGENDAMENTO')) &&
    row.some((cell) => normalize(cell).includes('STATUS DA CARGA')));
  if (header < 0) throw new Error('Cabeçalho do Report não reconhecido.');
  const headers = values[header].map((value) => normalize(value));
  const definitions: { [key: string]: string[] } = {
    scheduled: ['AGENDAMENTO (DATA/HORA)', 'AGENDAMENTO'],
    arrival: ['APRESENTAÇÃO DO VEÍCULO (DATA/HORA)', 'APRESENTACAO DO VEICULO'],
    operation: ['TIPO DE OPERAÇÃO'], quantity: ['QTD. (SACA/KG)', 'QTD. SACA/KG', 'QTD'],
    pallets: ['QTDE. PALETES', 'QTDE PALETES'], load: ['N° CARGA', 'N CARGA'],
    driver: ['MOTORISTA'], plate: ['PLACA'], trailer: ['PLACA DAS CARRETAS', 'CARRETA'],
    carrier: ['TRANSP.', 'TRANSPORTADORA'], vehicleStatus: ['STATUS VEÍCULO'],
    loadStatus: ['STATUS DA CARGA'], billingStatus: ['STATUS DO FATURAMENTO'],
    observation: ['OBSERVAÇÃO'], operationEnd: ['TÉRMINO DA OPERAÇÃO (DATA/HORA)', 'TERMINO DA OPERACAO'],
    statusAux: ['STATUS'], indicator: ['IND'],
    loadingStart: ['INÍCIO (DATA & HORA)', 'INÍCIO DATA & HORA'],
    loadingEnd: ['TÉRMINO (DATA & HORA)', 'TÉRMINO DATA & HORA'],
    loadingTime: ['TEMPO DE CARREGAMENTO'], palletObs: ['OBS. PALETE (100X100 | 100X120)', 'OBS. PALETE']
  };
  const fields = Object.keys(definitions);
  const columns = fields.map((field) => {
    const candidates = definitions[field].map((value) => normalize(value));
    if (field === 'statusAux') return headers.findIndex((cell, index) => cell === 'STATUS' && index + startColumn > 10);
    let index = headers.findIndex((cell) => candidates.includes(cell));
    if (index < 0) index = headers.findIndex((cell) => cell !== '' && candidates.some((candidate) => cell.includes(candidate) || candidate.includes(cell)));
    return index;
  });
  ['scheduled', 'load', 'operation', 'loadStatus'].forEach((field) => {
    if (columns[fields.indexOf(field)] < 0) throw new Error('Coluna obrigatória ausente: ' + field);
  });
  const strings: string[] = [];
  const dictionary = new Map<string, number>();
  const rows: (number | boolean | null | number[])[][] = [];
  for (let index = header + 1; index < values.length; index++) {
    const row = values[index];
    const load = row[columns[fields.indexOf('load')]];
    const scheduled = row[columns[fields.indexOf('scheduled')]];
    if ((load === '' || load === undefined) && (scheduled === '' || scheduled === undefined)) continue;
    const encoded: (number | boolean | null | number[])[] = [startRow + index + 1];
    columns.forEach((column) => {
      const value = column < 0 ? '' : row[column];
      if (value === '' || value === undefined || /^#(REF!|DIV\/0!|VALUE!|NAME\?|N\/A|NUM!|NULL!|SPILL!|CALC!)/.test(String(value))) {
        encoded.push(null);
      } else if (typeof value === 'string') {
        let key = dictionary.get(value);
        if (key === undefined) { key = strings.length; strings.push(value); dictionary.set(value, key); }
        encoded.push([key]);
      } else encoded.push(value);
    });
    rows.push(encoded);
  }
  if (!rows.length) throw new Error('Report sem registros válidos. A base publicada será preservada.');
  const result = JSON.stringify({ format: 'valorem-report-v1', sourceFile: workbook.getName(), sheet: 'Report',
    sourceModifiedAt: modifiedAt || new Date().toISOString(), fields: ['row', ...fields], strings, rows });
  // repository_dispatch accepts 64 KiB. Measure the actual JSON string wrapper too.
  const wrapped = JSON.stringify({ data: result });
  let bytes = 0;
  for (const char of wrapped) {
    const code = char.codePointAt(0) || 0;
    bytes += code <= 0x7f ? 1 : code <= 0x7ff ? 2 : code <= 0xffff ? 3 : 4;
  }
  if (bytes > 62000) throw new Error('A base excedeu o limite do conector GitHub (' + bytes + ' bytes). Nenhum registro foi truncado; a publicação anterior foi preservada.');
  console.log('Report preparado: ' + rows.length + ' registros, ' + bytes + ' bytes.');
  return result;
}
