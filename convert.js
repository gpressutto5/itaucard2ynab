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
const allData = XLSX.utils.sheet_to_json(
  workbook.Sheets[workbook.SheetNames[0]],
  { header: 1 }
);
const csvData = allData
  .filter(validateRow)
  .map(maybeHandleInstalments)
  .filter((row) => row !== null);
csvData.unshift(["Date", "Payee", "", "Outflow"]);

// push international transactions
const internationalStart = allData.findIndex(
  (row) => row[0] === "lançamentos internacionais"
);
const internationalEnd = allData.findIndex(
  (row) =>
    row[0] === "total de lançamentos internacionais (titular + adicionais)"
);
if (internationalStart !== -1 && internationalEnd !== -1) {
  const internationalSection = allData.slice(
    internationalStart + 1,
    internationalEnd
  );

  const internationalTransactions = [];
  let iofIndexes = internationalSection
    .map((row, index) =>
      row[0] === "IOF - transação internacional" ? Number(index) : null
    )
    .filter(Number);
  const internationSectionParts = iofIndexes.map((iofIndex, index) =>
    internationalSection.slice(
      Number(iofIndexes[index - 1] || -1) + 1,
      iofIndex + 1
    )
  );

  internationSectionParts.forEach((part) => {
    const partTransactions = [];
    for (let i = 0; i < part.length; i++) {
      const row = part[i];
      if (String(row[0]).match(/^\d{2}\/\d{2}\/\d{4}$/)) {
        const date = row[0];
        const payee = part[i + 1][1];
        const outflow = part[i + 1][3];
        partTransactions.push([date, payee, "", outflow]);
      }
    }
    const iof = part.find(
      (row) => row[0] === "IOF - transação internacional"
    )[3];

    // distribute iof among transactions by percentage
    const totalInternationalOutflow = partTransactions.reduce(
      (acc, row) => acc + row[3],
      0
    );
    partTransactions.forEach((row) => {
      const percentage = row[3] / totalInternationalOutflow;
      row[3] = Number((row[3] + iof * percentage).toFixed(2));
    });
    internationalTransactions.push(...partTransactions);
  });

  csvData.push(...internationalTransactions);
}

const csv = XLSX.utils.sheet_to_csv(XLSX.utils.aoa_to_sheet(csvData));
console.log(`Writing ${outputPath}`);
fs.writeFileSync(outputPath, csv);
