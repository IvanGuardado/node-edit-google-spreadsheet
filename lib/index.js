"use strict";

//module for using the google api to get anayltics data in an object
require("colors");
var request = require("request");
var _ = require("lodash");
var util = require("./util");
var Metadata = require('./metadata');
var async = require('async');


var Spreadsheet = require("./spreadsheet");
var Documents = require("./documents");

//public api
exports.create = function(opts, callback) {

  //validate options
  if(!opts)
    throw "Missing options";
  if(typeof opts.callback === 'function')
    callback = opts.callback;
  if(!callback)
    throw "Missing callback";
  if(!(opts.username && opts.password) && !opts.oauth)
    return callback("Missing authentication information");
  if(!opts.spreadsheetId  && !opts.spreadsheetName)
    return callback("Missing 'spreadsheetId' or 'spreadsheetName'");
  if(!opts.worksheetId  && !opts.worksheetName)
    return callback("Missing 'worksheetId' or 'worksheetName'");
  // if(opts.createNew && opts.worksheetName !== 'Sheet 1')
  //   return callback("Worksheet must be named Sheet 1 when creating new Spreadsheet");

  //default to http's' when undefined
  opts.useHTTPS = opts.useHTTPS === false ? '' : 's';
  
  var steps = [];

  if(opts.spreadsheetCreate) {
    var documents = new Documents(opts);
    steps.push(documents.auth);
    steps.push(documents.createSheet);
  }

  var spreadsheet = new Spreadsheet(opts);
  steps.push(spreadsheet.auth);
  steps.push(spreadsheet.init);

  async.series(steps, callback);


  // if(err && err.match(/spread(.*?)not found$/) && opts.createNew) {

};

Spreadsheet.prototype.getNames = function(curr) {
  var _this = this;
  return curr.val
    .replace(/\{\{\s*([\-\w\s]*?)\s*\}\}/g, function(str, name) {
      var link = _this.names[name];
      if(!link) return _this.log(("WARNING: could not find: " + name).yellow);
      return util.int2cell(link.row, link.col);
    })
    .replace(/\{\{\s*([\-\d]+)\s*,\s*([\-\d]+)\s*\}\}/g, function(both,r,c) {
      return util.int2cell(curr.row + util.num(r), curr.col + util.num(c));
    });
};

Spreadsheet.prototype.addVal = function(val, row, col) {

  // _this.log(("Add Value at R"+row+"C"+col+": " + val).white);

  if(!this.entries[row]) this.entries[row] = {};
  if(this.entries[row][col])
    this.log(("WARNING: R"+row+"C"+col+" already exists").yellow);

  var obj = { row: row, col: col },
      t = typeof val;
  if(t === 'string' || t === 'number')
    obj.val = val;
  else
    obj = _.extend(obj, val);

  if(obj.name)
    if(this.names[obj.name])
      throw "Name already exists: " + obj.name;
    else
      this.names[obj.name] = obj;

  if(obj.val === undefined && !obj.ref)
    this.log(("WARNING: Missing value in: " + JSON.stringify(obj)).yellow);

  this.entries[row][col] = obj;
};

Spreadsheet.prototype.compile = function() {

  var row, col, strs = [];
  this.maxRow = 0;
  this.maxCol = 0;
  for(row in this.entries)
    for(col in this.entries[row]) {
      var obj = this.entries[row][col];
      this.maxRow = Math.max(this.maxRow, row);
      this.maxCol = Math.max(this.maxCol, col);
      if(typeof obj.val === 'string')
        obj.val = this.getNames(obj);

      if(obj.val === undefined)
        continue;
      else
        obj.val = _.escape(obj.val.toString());

      strs.push(this.entryTemplate(obj));
    }

  return strs.join('\n');
};

Spreadsheet.prototype.send = function(options, callback) {
  if(typeof options === 'function')
    callback = options;
  else if(!callback)
    callback = function() {};

  
  if(!this.token)
    return callback("No authorization token. Use auth() first.");
  if(!this.bodyTemplate || !this.entryTemplate)
    return callback("No templates have been created. Use setTemplates() first.");

  var _this = this,
      entries = this.compile(),
      body = this.bodyTemplate({ entries: entries });

  //finally send all the entries
  _this.log(("Updating Google Docs...").grey);
  // _this.log(entries.white);
  async.series([
    function autoSize(next){
      if(!options || !options.autoSize)
        return next();
      _this.metadata(function(err, metadata){
        if(err) return next(err);

        //no resize needed
        if(metadata.rowCount >= _this.maxRow &&
           metadata.colCount >= _this.maxCol){
          return next(null);
        }

        //resize with maximums
        metadata.rowCount = Math.max(metadata.rowCount, _this.maxRow);
        metadata.colCount = Math.max(metadata.colCount, _this.maxCol);
        _this.metadata(metadata, next);
      });
    },
    function send(next){
      request({
        method: 'POST',
        url: _this.baseUrl() + '/batch',
        headers: _this.authHeaders,
        body: body
      }, function(err, response, body) {
        if(err) return next(err);
        if(body.indexOf("success='0'") >= 0) {
          err = "Error Updating Spreadsheet";
          _this.log(err.red.underline + ("\nResponse:\n" + body));
        } else {
          _this.log("Successfully Updated Spreadsheet".green);
          //data has been successfully sent, clear it
          _this.reset();
        }
        next(err);
      });
    }
  ], callback);
};

Spreadsheet.prototype.receive = function(callback) {
  if(!this.token)
    return callback("No authorization token. Use auth() first.");

  var _this = this;
  // get whole spreadsheet
  request({
    method: 'GET',
    url: this.baseUrl()+'?alt=json',
    headers: this.authHeaders
  }, function(err, response, body) {
    
    //body is error
    if(response.statusCode != 200)
      err = ''+body;

    //show error
    if(err)
      return callback(err, null);

    var result;
    try {
      result = JSON.parse(body);
    } catch(e) {
      return callback("JSON Parse Error: " + e);
    }

    if(!result.feed) {
      err = "Error Reading Spreadsheet";
      _this.log(
        err.red.underline +
        ("\nData:\n") + JSON.stringify(this.entries, null, 2) +
        ("\nResponse:\n" + body));
      callback(err, null);
      return;
    }

    var entries = result.feed.entry || [];
    var rows = {};
    var info = {
      spreadsheetId: _this.spreadsheetId,
      worksheetId: _this.worksheetId,
      worksheetTitle: result.feed.title.$t || null,
      worksheetUpdated: result.feed.updated.$t || null,
      authors: result.feed.author && result.feed.author.map(function(author) {
        return { name: author.name.$t, email: author.email.$t };
      }),
      totalCells: entries.length,
      totalRows: 0,
      lastRow: 1,
      nextRow: 1
    };
    var maxRow = 0;

    _.each(entries, function(entry) {

      var cell = entry.gs$cell,
          r = cell.row, c = cell.col;

      if(!rows[r])
        info.totalRows++, rows[r] = {};

      rows[r][c] = util.gcell2cell(cell);
      info.lastRow =  util.num(r);
    });

    if(entries.length)
      info.nextRow = info.lastRow+1;

    _this.log(("Retrieved "+entries.length +" cells and "+info.totalRows+" rows").green);

    callback(null,rows,info);


  });
};

Spreadsheet.prototype.metadata = function(data, callback){
  var meta = new Metadata(this);
  if(typeof data === 'function') {
    callback = data;
    meta.get(callback);
    return;
  } else if(!callback){ 
    callback = function() {};
  }
  meta.set(data, callback);
  return;
};

