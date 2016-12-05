/*
 * main.js
 */
'use strict';


const file = 'payment-terms.xlsx';

document.addEventListener('DOMContentLoaded', main);

function main() {
  fetch(file)
  .then(response => response.arrayBuffer())
  .then(arrayBuffer => {
    const data = new Uint8Array(arrayBuffer);

    const binaryString = data.reduce(
      (acc, cur) => acc + String.fromCharCode(cur), '')

    const workbook = XLSX.read(binaryString, {type: 'binary'});
    const terms = sheetToJSON(workbook.Sheets.Sheet1);

  })
}

function sheetToJSON(sheet) {
  const A = 65;
  const columnNames = [];
  const result = [];

  let column = A;
  let row = 1;
  let cell = String.fromCharCode(column) + row;

  do {
    columnNames.push(sheet[cell].v)
    column++
    cell = String.fromCharCode(column) + row;
  } while (cell in sheet)

  column = A;
  row = 2;
  cell = String.fromCharCode(column) + row;
  do {
    let record = {};

    do {
      let key = columnNames[column - A];
      let value = sheet[cell].v;
      record[key] = value;

      column++
      cell = String.fromCharCode(column) + row;
    } while (cell in sheet)

    result.push(record);

    column = A;
    row++;
    cell = String.fromCharCode(column) + row;
  } while (cell in sheet)

  return result;
}
