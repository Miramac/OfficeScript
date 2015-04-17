/**
 *  OfficeScript.XLS.NET.js
 *
 *  Connection C#/JavaScript for all global (scope: "presentation") Functions
 **/
/**console.log($range('a1', 'Sheet2').val(23).val());
* @class XLSNET
**/
var XLSNET = {
   
    findRange: function (rangeSelector, sheet) {
        rangeSelector = rangeSelector.split(',');
        return XLS.FindRange(rangeSelector, sheet);
    }
};

module.exports = XLSNET;