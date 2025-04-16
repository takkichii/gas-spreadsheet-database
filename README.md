# Use spreadsheet as database with Google Apps Script
- You can use the crud operation with spreadsheet
- Set the ID of a default spreadsheet to the Config.gs file

## Spreadsheet Sample
![Sample](img-sample-spreadsheet.png)

## Get Data from Spreadsheet

```javascript
const mdl = new Model();

const conditions = [
  { key: "ColumnName1", value: "value1" },
  { key: "ColumnName2", value: "value2" },
  // etc.
];

const offset = 10;
const limit = 5;

const allData         = mdl.getData("SheetName");
const dataWithAnd     = mdl.getData("SheetName", conditions);
const dataWithOr      = mdl.orGetData("SheetName", conditions);
const dataOffsetLimit = mdl.getData("SheetName", null, offset, limit);

Logger.log(allData);
Logger.log(dataWithAnd);
Logger.log(dataWithOr);
Logger.log(dataOffsetLimit);
```

### Use other spreadsheets

```javascript
const mdl = new Model("9Lnk1lZhQ9Hr8IMt60FIbDwLSlv5jlu9");
```

### Insert Data into Spreadsheet

```javascript
const mdl = new Model();

const keyValuePairs = [
  {
    "ColumnName1": "value1-1",
    "ColumnName2": "value2-1",
    // etc.
  },
  {
    "ColumnName1": "value1-2",
    "ColumnName2": "value2-2",
    // etc.
  },
  // etc.
];

const result = mdl.inserData("SheetName", keyValuePairs);

Logger.log(result);
```

## Update Data of Spreadsheet

```javascript
const mdl = new Model();

const conditions = [
  { key: "ColumnName1", value: "value1" },
  { key: "ColumnName2", value: "value2" },
  // etc.
];

const keyValuePair = {
  "ColumnName1": "new value 1",
  "ColumnName2": "new value 2",
  // etc.
};

const resultOfUpdate   = mdl.updateData("SheetName", keyValuePair, conditions);
const resultOfOrUpdate = mdl.orUpdateData("SheetName", keyValuePair, conditions);

Logger.log(resultOfUpdate);
Logger.log(resultOfOrUpdate);
```

## Delete Data of Spreadsheet

```javascript
const mdl = new Model();

const conditions = [
  { key: "ColumnName1", value: "value1" },
  { key: "ColumnName2", value: "value2" },
  // etc.
];

const resultOfDelete   = mdl.deleteData("SheetName", conditions);
const resultOfOrDelete = mdl.orDeleteData("SheetName", conditions);

Logger.log(resultOfDelete);
Logger.log(resultOfOrDelete);
```

## Truncate Data of Spreadsheet
- Delete all rows of the selected sheet except headers

```javascript
const mdl = new Model();

const result = mdl.truncateData("SheetName");

Logger.log(result);
```