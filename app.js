var express = require('express');
var path = require('path');
var bodyParser = require('body-parser');
var app = express();

app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());

// Add headers
app.use(function (req, res, next) {

  // Website you wish to allow to connect
  res.setHeader('Access-Control-Allow-Origin', '*');

  // Request methods you wish to allow
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS, PUT, PATCH, DELETE');

  // Request headers you wish to allow
  res.setHeader('Access-Control-Allow-Headers', 'X-Requested-With,content-type');

  // Set to true if you need the website to include cookies in the requests sent
  // to the API (e.g. in case you use sessions)
  res.setHeader('Access-Control-Allow-Credentials', true);

  // Pass to next layer of middleware
  next();
});

app.post('/openWord',function(req,res,next)
{
  var jsDavUrl=req.body.jsDavUrl;
  var wordPath=req.body.wordPath;
  var docPath=req.body.docPath;
  console.log(req.body);
  var exec = require('child_process').exec;
  var command=wordPath+"\""+ " \""+jsDavUrl+docPath;
  var cmd = "\""+command+"\"";
  //var cmd = '"C:/Program Files (x86)/Microsoft Office/root/Office16/WINWORD.EXE" "http://localhost:8000/app4office.docx"';

  exec(cmd, function(error, stdout, stderr) {
    if(error)
      console.error(error);
    res.end();
  });

});

// catch 404 and forward to error handler
app.use(function(req, res, next) {
  var err = new Error('Not Found');
  err.status = 404;
  next(err);
});

module.exports = app;
