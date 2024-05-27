# Spreadsheet-BDL

This is an attempt to put a behavioural definition language, inspired by the likes of [Gherkin](https://en.wikipedia.org/wiki/Cucumber_(software)#Gherkin_language), around an Excel spredsheet. This allows a set of processing tasks to be defined in human terms, and applied to a whole spreasheet. This is intended for cases where a huge ammount of data exists in a spreadsheet, and bulk processing is required which cannot be easily expressed using standard spreadsheet formulas. This is a proof of concept/experiement based on a problem my girlfriend described to me she was having.

See LICENSE file for license/usage details.

## Getting started

This is a node.js project, and was written against `v20.13.1`. Node.js can be downloaded [here](https://nodejs.org/en/download/prebuilt-installer).

Once node.js is installed, clone or copy this repo, and run `npm install` in the downloaded directotry in a command/terminal window to pull in any required dependancies.

Once done, the tool can be run using the following command:

```
node src/index.js <config file location> <spreadsheet file location> [debug]
```

where:

- config file location is the filepath to a configuration file containing the set of instructions to run
- spreadsheet file location is the filepath to the spreadsheet to process
- debug is an optional argument to turn on more duebug output

When run, the tool generates a new spreadsheet next to the provided starting file with `.processed.xslx` at the end of the file name.

This tool uses [exceljs](https://www.npmjs.com/package/exceljs) to interact with the spreadsheet.

### Available instructions

Available commands are as follows. Commands are case and character sensitive - they must exactly match to be processed. Commands need to exist on their own line/be new line seperated. 

- `Given spreadsheet sheet "X"` - selects/declares a particular sheet in the spreadsheet. Allow following commands will interact within that sheet.
- `Given column headings are on row "X"` - declares what row column headings are on. Any action which works on data will add one to this value (ie the first row of data following any column headings)
- `Given "X" is in column "Y"` - declares a column heading name (X), and which column it is in (Y). Any later command can then refer to the name (X) for processing
- `Create a duplicate sheet "X"` - creates a duplicate worksheet (called X) from the current worksheet. All data from the current worksheet is copied. The new worksheet is where any following commands will take effect.
- `Create a new column heading "X" in column "Y"` - creates and declares a new column. The name is provided via X, and which column it is in by Y. Any later command can then refer to the name (X) for processing.
- `If the "X" column for a row contains any of the following words "Y", set the "A" to "B"` - When processing data, if the named column X contains a set of words (Y - which can be comma seperated, eg `foo,bar`), set the cell value for that row and named column A to literal value B. Use this to set a literal value conditionally.
- `If the "X" column for a row contains any of the following words "Y", multiply the "A" by "B" and set the result in the "C" column` - When processing data, if the named column X contains a set of words (Y - which can be comma seperated, eg `foo,bar`), multiply the named column A by the value given in B, with the result being set in named column C. Use this to conditionally derive values.

When run, a context of instructions is maintained by the tool - so instructions can build upon each other. Eg, one instrunction can create a column by name, and a later instruction can put data into it by using that named column.

An example config file and spreasheet are in the 'examples' folder, and can be run as follows once cloned:

```
node src/index.js ./example/sampleConfig.txt ./example/sampleSpreadsheet.xlsx
```
