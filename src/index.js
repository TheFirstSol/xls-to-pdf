const path = require("path");
const http = require("http");
const express = require("express");
const cors = require("cors");
const multer = require("multer");
const upload = multer({ dest: "uploads/" });
const port = process.env.port || 3000;

const app = express();
app.use(express.json());
app.use(cors());
app.options("*", cors());

const server = http.createServer(app);

server.listen(port, () => {
  console.log(`server is listening on port: ${port}`);
});

app.get("/", (req, res) => {
  res.send("Welcome to xlstopdf");
});

app.post("/", upload.single("excel"), (req, res) => {
  console.log("asd", req.file, req.files);
  res.send("file recieved");
});
