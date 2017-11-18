require('date-utils');
var express = require('express');
var path = require('path');
var favicon = require('serve-favicon');
var logger = require('morgan');
var cookieParser = require('cookie-parser');
var bodyParser = require('body-parser');

var fs = require('fs'); // Add
var multer = require('multer'); // Add

var index = require('./routes/index');
var users = require('./routes/users');

var app = express();

// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'jade');

// uncomment after placing your favicon in /public
//app.use(favicon(path.join(__dirname, 'public', 'favicon.ico')));
app.use(logger('dev'));
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: false }));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

// app.use('/', index);

// Add start

//app.get('/', function(req, res){
//  fs.readdir('./files/uploads', function(error, list) {
//    if (error) {
//      // TODO:
//    } else {
//   // './files/conversion' も render できるようにする
//      // おそらく、それぞれのリストのコールバックを作成する感じ
//      res.render('index', {'upFiles' : list});
//   }
//  });
//});

// 非同期はあきらめて
var upList = function (req, res, next) {
  req.upList = fs.readdirSync('./files/uploads');
  next();
};
var convList = function (req, res, next) {
  req.convList = fs.readdirSync('./files/conversion');
  next();
};

app.use(upList);
app.use(convList);

app.get('/', function(req, res){
  res.render('index', {'upFiles' : req.upList, 'convFiles' : req.convList});
});

var upload = multer({ storage:
  multer.diskStorage({
    destination: function (req, file, cb) {
      cb(null, './files/uploads');
    },
    filename: function (req, file, cb) {
      var dt = new Date();
      var formatted = dt.toFormat("YYYYMMDD_HH24MISS");
      cb(null, '[' + formatted + ']' + file.originalname);
    }
  })
});

app.post('/', upload.single('uploadedfile'), function (req, res) {
  console.log(req.file);
  res.redirect(301, '/');
});

app.get('/upFiles/:file', function(req, res){
  fs.readFile('./files/uploads/' + req.params.file,
    function(err, data) {
    res.send(data, 200);
  });
});

app.get('/convFiles/:file', function(req, res){
  fs.readFile('./files/conversion/' + req.params.file,
    function(err, data) {
    res.send(data, 200);
  });
});

// Add end

app.use('/users', users);

// catch 404 and forward to error handler
app.use(function(req, res, next) {
  var err = new Error('Not Found');
  err.status = 404;
  next(err);
});

// error handler
app.use(function(err, req, res, next) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get('env') === 'development' ? err : {};

  // render the error page
  res.status(err.status || 500);
  res.render('error');
});

module.exports = app;
