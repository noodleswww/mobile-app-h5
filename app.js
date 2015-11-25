var express = require('express');
var path = require('path');
var favicon = require('serve-favicon');
var logger = require('morgan');
var cookieParser = require('cookie-parser');
var bodyParser = require('body-parser');

var routes = require('./routes/index');
var users = require('./routes/users');

var xlsx = require('node-xlsx');
var fs = require('fs');
var app = express();

// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');

// uncomment after placing your favicon in /public
//app.use(favicon(path.join(__dirname, 'public', 'favicon.ico')));
app.use(logger('dev'));
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({extended: false}));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));
app.get('/wx', function (req, res) {
  res.sendFile(path.join(__dirname + '/views/share-wx.html'));
});
app.get('/app', function (req, res) {
  res.sendFile(path.join(__dirname + '/views/share-app.html'));
});
app.get('/detail', function (req, res) {
  res.sendFile(path.join(__dirname + '/views/detail.html'));
});
app.get('/wenda', function (req, res) {
  res.sendFile(path.join(__dirname + '/views/wenda.html'));
});
app.get('/pie', function (req, res) {
  res.sendFile(path.join(__dirname + '/views/test.html'));
});
app.get('/tips', function (req, res) {
  res.sendFile(path.join(__dirname + '/views/wx-tips.html'));
});
app.get('/excel', function (req, res, next) {
  var data = [
    [1, 2, 3],      //用于设置表头
    [true, false, null, 'sheetjs'],   //内容区域
    ['foo', 'bar', new Date('2014-02-19T14:30Z'), '0.3'],
    ['baz', null, 'qux']
  ];
  var buffer = xlsx.build([{name: "students", data: data},{name:"students2",data:data}]); // returns a buffer
  fs.writeFileSync('excel.xlsx',buffer,'utf8');
  res.status(200).sendFile(__dirname+'/excel.xlsx');
});

// catch 404 and forward to error handler
app.use(function (req, res, next) {
  var err = new Error('Not Found');
  err.status = 404;
  next(err);
});

// error handlers

// development error handler
// will print stacktrace
if (app.get('env') === 'development') {
  app.use(function (err, req, res, next) {
    res.status(err.status || 500);
    res.render('error', {
      message: err.message,
      error: err
    });
  });
}

// production error handler
// no stacktraces leaked to user
app.use(function (err, req, res, next) {
  res.status(err.status || 500);
  res.render('error', {
    message: err.message,
    error: {}
  });
});


module.exports = app;
