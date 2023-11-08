import Excel from 'exceljs';
import neo4j, { Session } from 'neo4j-driver';
import config from './config';

export const driver = neo4j.driver(
  config.neo4j.connectionUri,
  neo4j.auth.basic(config.neo4j.username, config.neo4j.password),
  { disableLosslessIntegers: true }
);

function camelCase(str: string) {
  return str
    .replace(/[\s-_]\s(.)/g, function ($1) { return $1.toUpperCase(); })
    .replace(/[\s-_]/g, '')
    .replace(/^(.)/, function ($1) { return $1.toUpperCase(); });
}

function lowerCamelCase(str: string) {
  return camelCase(str)
    .replace(/^(.)/, function ($1) { return $1.toLowerCase(); });
}

async function addObjects(objects: any, properties: string[], labels: string | string[], modelName: string, session: Session) {
  if (!Array.isArray(labels))
      labels = [labels];
  labels.unshift('NeoCore__Object');
  let query = `
      UNWIND $objects AS object
      CREATE (n:${labels.join(':')} {enamespace: $modelName})
  `;
  if (properties && Array.isArray(properties) && properties.length)
      query += 'SET ' + properties.map(p => `n.${p} = object.${p}`).join(',');
  return session.run(query, { modelName, objects });
}

let modelName = 'MiroModel_uXjVM0-Zopo=';
let metamodelName = 'Workbook';

const workbook = new Excel.Workbook();
workbook.xlsx.readFile('./example.xlsx').then(async workbook => {
  let sheets : string[] = [];
  let sheetRows : Array<Array<Record<string, string | boolean | number>>> = [];
  let sheetColumns: string[][] = [];
  workbook.eachSheet(function (worksheet) {
    let header = worksheet.getRow(1);
    let columnCount = 1;
    let columnNames: string[] = [];
    let rows: Array<Record<string, string | boolean | number>> = [];
    let cell: Excel.Cell;
    do {
      cell = header.getCell(columnCount++);
      if (cell.text)
        columnNames.push(lowerCamelCase(cell.text))
    } while (cell.text)
    columnCount = columnNames.length;
    if (columnCount == 0) {
      throw new Error("No column names exist");
    }
    let rowIndex = 2;
    let row = worksheet.getRow(rowIndex++);
    while(row.getCell(1).text) {
      let rowMap: Record<string,string|boolean|number> = {};
      for(let i = 0; i < columnCount; i++) {
        let cell = row.getCell(i+1);
        rowMap[columnNames[i]] = cell.value as string|boolean|number;
      }
      rows['enamespace'] = modelName;
      rows.push(rowMap);
      row = worksheet.getRow(rowIndex++)
    }
    sheets.push(worksheet.name)
    sheetColumns.push(columnNames);
    sheetRows.push(rows)
  });
  let session = driver.session();
  for (let i = 0; i < sheets.length; i++) {
    let sheet = sheets[i];
    let rows = sheetRows[i];
    let sheetRowType = 'Spreadsheet__' + camelCase(sheet) + 'Row';
    let labels = ['Spreadsheet__Row', sheetRowType];
    await addObjects(rows,sheetColumns[i],labels,modelName,session);
    await session.run(`CREATE (s:NeoCore__Object:Spreadsheet__Sheet {enamespace:$modelName, name:$sheet}) WITH s MATCH (r:${sheetRowType} {enamespace:$modelName}) CREATE (s)-[:rows]->(r)`,{modelName, sheet});
  }
  await session.run(`CREATE (w:NeoCore__Object:Spreadsheet__Workbook {enamespace:$modelName}) WITH w MATCH (s:Spreadsheet__Sheet {enamespace:$modelName}) CREATE (w)-[:sheets]->(s)`,{modelName});
  await session.close();
  await driver.close();
});
