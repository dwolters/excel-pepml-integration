import commandLineArgs from 'command-line-args'
import commandLineUsage from 'command-line-usage'
import Excel from 'exceljs';
import neo4j, { Session } from 'neo4j-driver';
import config from './config';

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

async function extractModel(filepath: string, modelName: string, connectionUri: string, username: string, password: string, justMetamodel: boolean | undefined) {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(filepath)
  let sheets: string[] = [];
  let sheetRows: Array<Array<Record<string, string | boolean | number>>> = [];
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
    while (row.getCell(1).text) {
      let rowMap: Record<string, string | boolean | number> = {};
      for (let i = 0; i < columnCount; i++) {
        let cell = row.getCell(i + 1);
        rowMap[columnNames[i]] = cell.value as string | boolean | number;
      }
      rows['enamespace'] = modelName;
      rows.push(rowMap);
      row = worksheet.getRow(rowIndex++)
    }
    sheets.push(worksheet.name)
    sheetColumns.push(columnNames);
    sheetRows.push(rows)
  });
  if (justMetamodel) {
    printMetamodel(sheets,sheetColumns,sheetRows)
  } else {
    storeModel(sheets, sheetRows, sheetColumns, modelName, connectionUri, username, password);
  }
}

function printMetamodel(sheets: string[], sheetColumns: string[][], sheetRows: Array<Array<Record<string, string | boolean | number>>>) {
  let sheetDefinitions = '';
  for(let i = 0; i < sheets.length; i++) {
    let sheet = sheets[i];
    let rows = sheetRows[i];
    let columns = sheetColumns[i];
    let columnTypes: string[];
    if(rows.length > 0) {
      columnTypes = columns.map(column => typeof rows[0][column]);
    } else {
      columnTypes = columns.map(() => 'string');
    }
    sheetDefinitions += `  ${sheet}Row: Row {\n`;
    for(let c = 0; c < columns.length; c++) {
      sheetDefinitions += `    ${columns[c]}: ${columnTypes[c]}\n`;
    }
    sheetDefinitions += "  }\n";
  }
  const mm = `metamodel Spreadsheet {
  Workbook {
    <+>-sheets(0..*)->Sheet
  }
  Sheet {
    .name: EString
    <+>-rows(0..*)->Row
  }
  Row
${sheetDefinitions}  
}`;
  console.log(mm);
}

async function storeModel(sheets: string[], sheetRows: Array<Array<Record<string, string | boolean | number>>>, sheetColumns: string[][], modelName: string, connectionUri: string, username: string, password: string) {
  const driver = neo4j.driver(
    connectionUri,
    neo4j.auth.basic(username, password),
    { disableLosslessIntegers: true }
  );
  let session = driver.session();
  for (let i = 0; i < sheets.length; i++) {
    let sheet = sheets[i];
    let rows = sheetRows[i];
    let sheetRowType = 'Spreadsheet__' + camelCase(sheet) + 'Row';
    let labels = ['Spreadsheet__Row', sheetRowType];
    await addObjects(rows, sheetColumns[i], labels, modelName, session);
    await session.run(`CREATE (s:NeoCore__Object:Spreadsheet__Sheet {enamespace:$modelName, name:$sheet}) WITH s MATCH (r:${sheetRowType} {enamespace:$modelName}) CREATE (s)-[:rows]->(r)`, { modelName, sheet });
  }
  await session.run(`CREATE (w:NeoCore__Object:Spreadsheet__Workbook {enamespace:$modelName}) WITH w MATCH (s:Spreadsheet__Sheet {enamespace:$modelName}) CREATE (w)-[:sheets]->(s)`, { modelName });
  await session.close();
  await driver.close();
}

const optionDefinitions = [
  {
    name: 'help',
    alias: 'h',
    type: Boolean,
    description: 'Display this usage guide.'
  },
  {
    name: 'file',
    alias: 'f',
    type: String,
    description: 'The input file (*.xlsx) to process',
    typeLabel: '<file>'
  },
  {
    name: 'name',
    alias: 'n',
    type: String,
    description: 'Name of the model in Neo4j',
    typeLabel: '<modelName>'
  },
  {
    name: 'username',
    alias: 'u',
    type: String,
    description: 'Name of Neo4j User',
    typeLabel: '<username>'
  },
  {
    name: 'password',
    alias: 'p',
    type: String,
    description: 'Password of Neo4j User',
    typeLabel: '<password>'
  },
  {
    name: 'connectionUri',
    alias: 'c',
    type: String,
    description: 'Connection URI',
    typeLabel: '<connectionUri>'
  },
  {
    name: 'metamodel',
    alias: 'm',
    type: Boolean,
    description: 'If set only the metamodel for this file is generated.',
  }
]

const options = commandLineArgs(optionDefinitions)

if (options.help) {
  const usage = commandLineUsage([
    {
      header: 'Example Usage',
      content: 'node lib/index.js -f ./example.xlsx -m ExampleWorkbook'
    },
    {
      header: 'Options',
      optionList: optionDefinitions
    }
  ])
  console.log(usage)
} else {
  console.log(options)
  let connectionUri = options.connectionUri || config.neo4j.connectionUri;
  let username = options.username || config.neo4j.username;
  let password = options.password || config.neo4j.password;
  extractModel(options.file, options.name, connectionUri, username, password, options.metamodel);
}