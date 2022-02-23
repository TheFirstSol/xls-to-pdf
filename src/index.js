const path = require("path");
const http = require("http");
const express = require("express");
const cors = require("cors");
const fs = require('fs');
const AWS = require('aws-sdk')
const XLSX = require('xlsx')
const pdf = require('html-pdf')
const options = {
  format: 'A4'
}
const multer = require("multer");
const upload = multer({
  dest: "uploads/"
});


const spaceEndpoint = new AWS.Endpoint('sgp1.digitaloceanspaces.com');

const s3 = new AWS.S3({
  endpoint: spaceEndpoint,
  accessKeyId: '6JFG4MC6AFB53QQCVMF2',
  secretAccessKey: '/FedMxOJUxJoa6dN3NN8QMk6ExgPLTNSa0v+kTaZ7WA',
  region: 'sgp1'
})
const port = process.env.port || 3000;

const app = express();
app.use(express.json());
app.use(cors());
app.options("*", cors());


let htmlFile = '';

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
  excelToHtml(workbook.Sheets, res)

});




const excelToHtml = async (sheets, res) => {

  htmlFile = ''
  for (let sheet in sheets) {
    if (typeof sheet !== 'undefined') {
      htmlFile += '<table summary="" class="turntable">' + '\n' + '<thead>';
      for (let cell in sheets[sheet]) {

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

    pdf.create(htmlFile, options).toFile(`pdfs/${new Date().getTime()}.pdf`, (err, result) => {
      if (err) {
        console.log(err)
        return res
          .status(400)
          .json({
            message: err.message
          });
      } else {
        console.log(result)
        const newPath = result.filename.split('/');
        const filePath = `pdfs/${newPath[newPath.length - 1]}`
        console.log(filePath)
        uploadToBucket(filePath, res)
      }
    })
  }
}






const uploadToBucket = async (path, res) => {
  const file = fs.readFileSync(path);
  console.warn(file)
  const key = `${new Date().getTime()}.pdf`
  s3.putObject({
    Bucket: 'exceltopdf',
    Key: key,
    Body: file,
    ACL: 'public-read'
  }, (err, data) => {
    if (err) {
      return res
        .status(400)
        .json({
          message: err.message
        });
    }
    console.log("Your file has been uploaded successfully!", data);
    getSignedFile(key, res)
  })
}



const getSignedFile = async (key, res) => {
  const expireSeconds = 60 * 60 * 5

  const url = await s3.getSignedUrl('getObject', {
    Bucket: 'exceltopdf',
    Key: key,
    Expires: expireSeconds
  });
  res.send(url);
}