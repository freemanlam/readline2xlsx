const fs = require('fs');
const readline = require('readline');
const XLSX = require('xlsx');

const DATASOURCE = 'data-source.txt';
const OUTPUT_SHEET_NAME = 'Output';
const OUTPUT_FILE_NAME = 'output.xlsx';

// create instance of readline
// each instance is associated with single input stream
let rl = readline.createInterface({
  input: fs.createReadStream(DATASOURCE);
});

let line_no = 0;

const data = [];
let lastLine;

// event is emitted after each line
rl.on('line', line => {
  line_no++;
  if (line_no % 2 === 1) {
    const date = line.slice(0, 10);
    const eventName = line.slice(10, line.length);
    lastLine = {
      Date: date.split('.').join('-'),
      Event: eventName.trim()
    };
  } else {
    lastLine.Desciption = line;
    data.push(lastLine);
    console.log(lastLine);
  }
});

// end
rl.on('close', line => {
  console.log('Total lines : ' + line_no);

  const sheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workbook, sheet, OUTPUT_SHEET_NAME);
  XLSX.writeFile(workbook, OUTPUT_FILE_NAME);
});
