const commandLineArgs = require("command-line-args");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const optionDefinitions = [
  { name: "input", type: String },
  { name: "output", type: String },
];

const options = commandLineArgs(optionDefinitions);

if (!options.input) {
  console.error("Please specify an input file");
  process.exit(1);
}

const inputPath = path.join(process.cwd(), options.input);

let output = options.output;
if (output === undefined) {
  output = options.input;
  output = output.replace(/\.xlsx?$/i, "");
  output += ".csv";
}
const outputPath = path.join(process.cwd(), output);

const validateRow = (row) => {
  return (
    row.length === 4 &&
    String(row[0]).match(/^\d{2}\/\d{2}\/\d{4}$/) &&
    String(row[1]) !== "dólar de conversão"
  );
};
const maybeHandleInstalments = (row) => {
  const payee = String(row[1]);
  const matches = payee.match(/(\d{2})\/(\d{2})$/);
  if (matches) {
    if (matches[1] !== "01") {
      return null;
    }
    const instalments = Number(matches[2]);
    row[3] = Number((row[3] * instalments).toFixed(2));
    row[1] = payee.replace(/ *\d{2}\/\d{2}$/, "");
  }
  return row;
};
var workbook = XLSX.readFile(inputPath);
const data = XLSX.utils
  .sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 })
  .filter(validateRow)
  .map(maybeHandleInstalments)
  .filter((row) => row !== null);
data.unshift(["Date", "Payee", "", "Outflow"]);

const csv = XLSX.utils.sheet_to_csv(XLSX.utils.aoa_to_sheet(data));
console.log(`Writing ${outputPath}`);
fs.writeFileSync(outputPath, csv);
