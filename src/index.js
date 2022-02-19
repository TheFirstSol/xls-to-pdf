const path = require("path");
const http = require("http");
const express = require("express");
const cors = require("cors");
const fs = require('fs')
const XLSX = require('xlsx')
const pdf = require('html-pdf')
const options = {
  format: 'A4'
}
const multer = require("multer");
const upload = multer({
  dest: "uploads/"
});
const port = process.env.port || 3000;

const app = express();
app.use(express.json());
app.use(cors());
app.options("*", cors());


var htmlFile = '';

const server = http.createServer(app);

server.listen(port, () => {
  console.log(`server is listening on port: ${port}`);
});

app.get("/", (req, res) => {
  res.send("Welcome to xlstopdf");
});

app.post("/", upload.single("excel"), async (req, res) => {
  console.log("asd", req.file);
  const workbook = XLSX.readFile(req.file.path);
  excelToHtml(workbook.Sheets)

  res.send("file recieved");
});




const excelToHtml = async (sheets) => {

  for (let sheet in sheets) {
    if (typeof sheet !== 'undefined') {
      htmlFile += '<table summary="" class="turntable">' + '\n' + '<thead>';
      for (var cell in sheets[sheet]) {

        if (typeof sheets[sheet][cell].w !== 'undefined') {
          if (cell === 'A1') {
            htmlFile += '\n' + '<tr>' + '\n' + '<th>' + sheets[sheet][cell].w.replace('& ', '&amp;').replace('-', '&ndash;').replace('–', '&mdash;') + '</th>';
          } else {
            if (cell === 'A2') {
              htmlFile += '\n' + '</tr>' + '\n' + '</thead>' + '\n' + '<tr>' + '\n' + '<th>' + sheets[sheet][cell].w.replace('& ', '&amp;').replace('-', '&ndash;').replace('–', '&mdash;') + '</th>';
            } else {
              if (cell.slice(0, 1) === 'A') {
                htmlFile += '\n' + '</tr>' + '\n' + '<tr>' + '\n' + '<th>' + sheets[sheet][cell].w.replace('& ', '&amp;').replace('-', '&ndash;').replace('–', '&mdash;') + '</th>';
              } else {
                htmlFile += '\n' + '<td>' + sheets[sheet][cell].w.replace('& ', '&amp;').replace('-', '&ndash;').replace('–', '&mdash') + '</td>';
              }
            }
          }
        }
      }
      htmlFile += '\n' + '</tr>' + '\n' + '</table>' + '\n';
    }

    pdf.create(htmlFile, options).toFile(`${new Date().getTime()}.pdf`, (err, result) => {
      if (err) {
        console.log(err)
      } else {
        console.log(result)
      }
    })
  }
}