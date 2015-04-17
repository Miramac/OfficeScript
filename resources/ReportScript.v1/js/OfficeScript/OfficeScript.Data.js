/**
 *  OfficeScript.Report.js
 *  OfficeScript Core: Connects Selectors and global Functions for Slides and Shapes
 *
 * 
 **/
var $ = require('./OfficeScript.Core')
, XLSNET = require('./OfficeScript.Data.NET')
, Range = require('./OfficeScript.Data.Range')
;

/**
 * Description
 * @class DataScript
 * @param {} selector
 * @param {} context
 * @return DataScript
 */
var DataScript = function (selector, context) {
    var DataScript = {
        version: '0.1.0-dev'
    };
    return DataScript;
};
/**
* @class $presentation
*/
var Workbook = {
    /**
     * Description
     * @method save
     * @return ThisExpression
     */
    save: function () {
        XLSNET.save();
        return this;
    }
    ,
    /**
      * Description
      * @method saveAs
      * @param {} path
      * @param {} type
      * @param {} embedFonts
      * @return ThisExpression
      */
    saveAs: function (path, type, embedFonts) {
        XLSNET.saveAs(path, type, embedFonts);
        return this;
    }
    ,
    /**
      * Description
      * @method saveCopyAs
      * @param {} path
      * @param {} type
      * @param {} embedFonts
      * @return ThisExpression
      */
    saveCopyAs: function (path, type, embedFonts) {
        XLSNET.saveCopyAs(path, type, embedFonts);
        return this;
    }
    , name: XLSNET.name
    , path: XLSNET.path
};

/**
 * Description
 * @method ppt
 * @param {} selector
 * @param {} context
 * @return NewExpression
 */
$.fn.ppt = function (selector, context) {
    this.selector = (selector) ? selector : this.selector;
    this.context = (context) ? context : this.context;
    return new DataScript(this.selector, this.context);
};


$.extend(DataScript, { range: Range });
$.extend(DataScript, { workbook: Workbook });

module.exports = DataScript;