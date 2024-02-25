const express = require('express');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
const port = 3000;

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

app.use(express.static(path.join(__dirname, 'public')));

app.get('/getexcel', (req, res) => {
  const excelFilePath = 'List.xlsx';

  readExcel(excelFilePath, (data) => {
    res.render('index', { data });
  });
});

async function readExcel(filePath, callback) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(1);

    let output = '';

    worksheet.eachRow((row, rowNumber) => {
      output += "Row ${rowNumber}: ${row.getCell(1).value}, ${row.getCell(2).value}<br>";
    });

    callback(output);

  } catch (error) {
    console.error('Error reading Excel sheet:', error.message);
  }
}

app.listen(port, () => {
  console.log("Server running at http://localhost:${port}");
});