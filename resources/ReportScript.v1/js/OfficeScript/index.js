
var $ = require('./OfficeScript.Core')
, Report = require('./OfficeScript.Report')
, Data = require('./OfficeScript.Data')
;

$.extend({ ppt: Report });
$.extend({ xls: Data });
module.exports = $;