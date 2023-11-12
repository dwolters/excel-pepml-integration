# Excel to Neo4j Model
This project allows to export excel data as a NeoCore model to Neo4j.
NeoCore is a technique to represent models in a Neo4j database.
The generated model can then be transformed using Triple Graph Grammars.
See [eMoflon::Neo](https://github.com/eMoflon/emoflon-neo) for more details on NeoCore and for model transformation.

## Prior to Usage

Run the following commands prior to using this tool for exporting Excel data to Neo4j.

```
npm install
tsc
```

## Persisting database credentials

Database credentials can be passed as command-line arguments. Alternatively, they can be set in [config.ts](src/config.ts).

## Usage
Export Excel data to Neo4j:
```
node lib/index.js -f ./example.xlsx -n Workbook
```

Generation of a eMSL metamodel:
```
node lib/index.js -f ./example.xlsx -n Workbook -m
```

Further options are available to set the database credentials as command-line arguments. See:
```
node lib/index.js -h
```

## Limitations
This project is only intended for Excel Workbooks without formulas and merged cells.
Furthermore, the first row of every sheet must provide the names for each column.