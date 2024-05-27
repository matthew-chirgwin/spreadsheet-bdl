const ExcelJS = require("exceljs");
const { readFile } = require("fs/promises");

// modifed by arg below
let DEBUG = false;

// set of regexes to match/provide a DDL style way of describing/interacting with a spreadsheet
const SHEET_REGEX = /Given spreadsheet sheet "(.+)"/;
const COLUMN_HEADER_ROW_REGEX = /Given column headings are on row "(.+)"/;
const COLUMN_HEADER_REGEX = /Given "(.+)" is in column "(.+)"/;
const CREATE_DUPLICATE_SHEET = /Create a duplicate sheet "(.+)"/;
const CREATE_NEW_COLUMN = /Create a new column heading "(.+)" in column "(.+)"/;
const SET_VALUE_TOKEN =
  /If the "(.+)" column for a row contains any of the following words "(.+)", set the "(.+)" to "(.+)"/;
const MULTIPLY_IF_TOKEN =
  /If the "(.+)" column for a row contains any of the following words "(.+)", multiply the "(.+)" by "(.+)" and set the result in the "(.+)" column/;

// Helper, get match from a regex result array
const getMatch = (regexResult, matchCount = 1) =>
  regexResult.slice(1, 1 + matchCount);

// function run when a command matches the `SHEET_REGEX` regex
const sheetRegexFn = (configStr, context) => {
  logDebug(`${configStr}: Starting`);
  const { workbook } = context;
  const [name] = getMatch(SHEET_REGEX.exec(configStr));
  const foundWorksheet = workbook.getWorksheet(name);
  if (!foundWorksheet) {
    logErrorAndExit(`Worksheet '${name}' not found in Workbook.`);
  } else {
    log(`${configStr}: OK`);
  }
  return {
    ...context,
    worksheet: foundWorksheet,
  };
};
// function run when a command matches the `COLUMN_HEADER_ROW_REGEX` regex
const columnRowRegexFn = (configStr, context) => {
  logDebug(`${configStr}: Starting`);
  const { worksheet } = context;
  const [row] = getMatch(COLUMN_HEADER_ROW_REGEX.exec(configStr));
  const hasRow = worksheet.rowCount >= Number(row);
  if (!hasRow) {
    logErrorAndExit(
      `Worksheet '${worksheet.getName()}' only has '${
        worksheet.rowCount
      }' rows. ${row} expected`
    );
  } else {
    log(`${configStr}: OK`);
  }
  return {
    ...context,
    columnRow: row,
  };
};
// function run when a command matches the `COLUMN_HEADER_REGEX` regex
const columnHeaderRegexFn = (configStr, context) => {
  logDebug(`${configStr}: Starting`);
  const { namedColumnData, columnRow, worksheet } = context;
  const [colName, cell] = getMatch(COLUMN_HEADER_REGEX.exec(configStr), 2);
  const cellLoc = `${cell}${columnRow}`;
  const valueInCell = worksheet.getCell(cellLoc).value;
  const columnIndex = worksheet.getCell(cellLoc).col;
  if (valueInCell !== colName) {
    logErrorAndExit(
      `Column heading '${colName}' is on row '${columnRow}'. Value found '${valueInCell}'`
    );
  } else {
    log(`${configStr}: OK`);
  }
  return {
    ...context,
    namedColumnData: [
      ...namedColumnData,
      { colName, cell, cellLoc, columnIndex },
    ],
  };
};
// function run when a command matches the `CREATE_DUPLICATE_SHEET` regex
const createDuplicateSheetRegexFn = (configStr, context) => {
  logDebug(`${configStr}: Starting`);
  const { workbook, worksheet } = context;
  const [name] = getMatch(CREATE_DUPLICATE_SHEET.exec(configStr));
  const currentWorksheetRows = worksheet
    .getRows(0, worksheet.rowCount)
    .map((row) => row.values);
  const newWorksheet = workbook.addWorksheet(`${name}`);
  newWorksheet.insertRows(0, currentWorksheetRows);
  log(`${configStr}: OK`);
  return {
    ...context,
    worksheet: newWorksheet,
  };
};
// function run when a command matches the `CREATE_NEW_COLUMN` regex
const createColumnRegexFn = (configStr, context) => {
  logDebug(`${configStr}: Starting`);
  const { namedColumnData, worksheet, columnRow } = context;
  const [colName, cell] = getMatch(CREATE_NEW_COLUMN.exec(configStr), 2);
  const cellLoc = `${cell}${columnRow}`;
  worksheet.getCell(cellLoc).value = colName;
  const columnIndex = worksheet.getCell(cellLoc).col;
  log(`${configStr}: OK`);
  return {
    ...context,
    namedColumnData: [
      ...namedColumnData,
      { colName, cell, cellLoc, columnIndex },
    ],
  };
};
// function run when a command matches the `SET_VALUE_TOKEN` regex
const setValueTokenRegex = (configStr, context) => {
  logDebug(`${configStr}: Starting`);
  const { namedColumnData, columnRow, worksheet } = context;
  const [colToCheck, tokens, colToSet, value] = getMatch(
    SET_VALUE_TOKEN.exec(configStr),
    4
  );

  const inspectColumn = namedColumnData.find(
    ({ colName }) => colName === colToCheck
  );
  const expectedTokens = tokens.split(",");
  const updateColumn = namedColumnData.find(
    ({ colName }) => colName === colToSet
  );

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > columnRow) {
      const dataToCheck = row.values[inspectColumn.columnIndex];
      logDebug(
        `Processing row ${rowNumber} - Column '${colToCheck}' row ${rowNumber} content: '${dataToCheck}'`
      );
      if (expectedTokens.includes(dataToCheck)) {
        logDebug(`Token matched`);
        const cellLoc = `${updateColumn.cell}${rowNumber}`;
        worksheet.getCell(cellLoc).value = value;
      }
    }
  });

  log(`${configStr}: OK`);
  return {
    ...context,
  };
};
// function run when a command matches the `MULTIPLY_IF_TOKEN` regex
const multiplyIfTokenRegexFn = (configStr, context) => {
  logDebug(`${configStr}: Starting`);
  const { namedColumnData, columnRow, worksheet } = context;
  const [colToCheck, tokens, colToMultiply, product, colToSet] = getMatch(
    MULTIPLY_IF_TOKEN.exec(configStr),
    5
  );

  const inspectColumn = namedColumnData.find(
    ({ colName }) => colName === colToCheck
  );
  const expectedTokens = tokens.split(",");
  const multiplyColumn = namedColumnData.find(
    ({ colName }) => colName === colToMultiply
  );
  const updateColumn = namedColumnData.find(
    ({ colName }) => colName === colToSet
  );

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > columnRow) {
      const dataToCheck = row.values[inspectColumn.columnIndex];
      logDebug(
        `Processing row ${rowNumber} - Column '${colToCheck}' row ${rowNumber} content: '${dataToCheck}'`
      );
      if (expectedTokens.includes(dataToCheck)) {
        logDebug(`Token matched`);
        const mpValue = row.values[multiplyColumn.columnIndex] * product;
        const cellLoc = `${updateColumn.cell}${rowNumber}`;
        worksheet.getCell(cellLoc).value = mpValue;
      }
    }
  });

  log(`${configStr}: OK`);
  return {
    ...context,
  };
};
// Map all regexes (as strings) to their raw regex and function for processing when iterating over instructions
const REGEX_FN_MAP = {
  [SHEET_REGEX]: {
    regex: SHEET_REGEX,
    fn: sheetRegexFn,
  },
  [COLUMN_HEADER_ROW_REGEX]: {
    regex: COLUMN_HEADER_ROW_REGEX,
    fn: columnRowRegexFn,
  },
  [COLUMN_HEADER_REGEX]: {
    regex: COLUMN_HEADER_REGEX,
    fn: columnHeaderRegexFn,
  },
  [CREATE_DUPLICATE_SHEET]: {
    regex: CREATE_DUPLICATE_SHEET,
    fn: createDuplicateSheetRegexFn,
  },
  [CREATE_NEW_COLUMN]: {
    regex: CREATE_NEW_COLUMN,
    fn: createColumnRegexFn,
  },
  [SET_VALUE_TOKEN]: {
    regex: SET_VALUE_TOKEN,
    fn: setValueTokenRegex,
  },
  [MULTIPLY_IF_TOKEN]: {
    regex: MULTIPLY_IF_TOKEN,
    fn: multiplyIfTokenRegexFn,
  },
};
// Get the set of all valid regexes for later processing
const ALL_VALID_CONFIG = Object.keys(REGEX_FN_MAP);

// Helper function for loging errors. Will exit the program when called.
const logErrorAndExit = (message) => {
  console.error(`ERROR: ${message} EXITING`);
  process.exit(1);
};

// general logging helpers. Debug varient only logs if the debug value is true
const log = console.log;
const logDebug = (message) => DEBUG && log(message);

// for a given instruction, confirm it matches a regex we have (and thuis is a valid instruction)
const regexMatchFound = (instruction) =>
  ALL_VALID_CONFIG.some((regexKey) => {
    const { regex } = REGEX_FN_MAP[regexKey];
    return regex.test(instruction);
  });

// take in a config file, iterate through the instructions, remove unrecognised instructions. Returns 
// a function which for a given workbook (Spredsheet) apllies those instructions. Instructions are provided
// with a context of prior actions/assertions made
const parseConfig = (configFile) => {
  const allInstructions = configFile.split("\n");
  logDebug(`Number of instructions in config: ${allInstructions.length}`);
  const validInstructions = allInstructions.filter((configLine) => {
    const matchFound = regexMatchFound(configLine);
    if (!matchFound) {
      logDebug(`Config instruction '${configLine}' not matched. Removing.`);
    }
    return matchFound;
  });
  logDebug(
    `Number of matched instructions in config: ${validInstructions.length}`
  );
  return (workbook) =>
    validInstructions.reduce(
      (context, instruction) => {
        return Object.entries(REGEX_FN_MAP)
          .filter(([, { regex }]) => instruction.match(regex) !== null)
          .map(([, { fn }]) => fn(instruction, context))[0];
      },
      {
        workbook,
        worksheet: undefined,
        columnRow: undefined,
        namedColumnData: [],
      }
    );
};

// run function - take input parameters, parses config, runs the commands and then writes the result
const run = async (configFilePath, sourceSpreadsheetPath) => {
  const workbook = new ExcelJS.Workbook();
  let runInstructions;
  try {
    await workbook.xlsx.readFile(sourceSpreadsheetPath);
  } catch (e) {
    logErrorAndExit(
      `Unable to read XLSX file at '${sourceSpreadsheetPath}': ${e.message}`
    );
  }
  try {
    runInstructions = parseConfig(
      await readFile(configFilePath, { encoding: "utf8" })
    );
  } catch (e) {
    logErrorAndExit(
      `Unable to read or parse config file at '${configFilePath}': ${e.message}`
    );
  }
  const newFileName = `${
    sourceSpreadsheetPath.split(".xlsx")[0]
  }.processed.xlsx`;
  log(
    `Processing Spreadsheet '${sourceSpreadsheetPath}'. Results produced to '${newFileName}'`
  );
  try {
    runInstructions(workbook);
  } catch (e) {
    logErrorAndExit(
      `Error running instructions: ${e.message}`
    );
  }
  try {
    await workbook.xlsx.writeFile(newFileName);
  } catch (e) {
    logErrorAndExit(
      `Unable to write XLSX file at '${newFileName}': ${e.message}`
    );
  }
  log("Done");
};

// get command args
const [configFileLocation, spreadsheetLocation, debug] = process.argv.slice(2);
DEBUG = debug !== undefined;

run(configFileLocation, spreadsheetLocation);
