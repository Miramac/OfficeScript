/**
 *  OfficeScript.Report.js
 *  OfficeScript Core: Connects Selectors and global Functions for Slides and Shapes
 **/
 
    var $ = require('./OfficeScript.Core')
    , Slides = require('./OfficeScript.Report.Slides')
    , Shapes = require('./OfficeScript.Report.Shapes')
    , PPTNET = require('./OfficeScript.PPT.NET')
    ;

/**
 * Description
 * @class ReportScript
 * @param {} selector
 * @param {} context
 * @return ReportScript
 */
var ReportScript = function (selector, context) {
    var ReportScript = {
        version: '0.3.2-dev'
    };
    return ReportScript;
};


/**
* @class $presentation
*/
var Presentation = {

    /**
    * Speichert die PowerPoint-Präsentation ab.
    * @method save
    * @chainable
    *
    * @example
    * Speichert die PowerPoint-Präsentation ab. 
    * @example
    *     $presentation.save();
    */
    save: function () {
        PPTNET.save();
        return this;
    },
    
    
    /**
    * Speichert die PowerPoint-Präsentation unter dem übergebenen Pfad unter dem übergebenen Namen und Typ ab.
    * WICHTIG: Diese Funktion erstellt keine fehlenden Ordner, der angegebene Pfad (bis auf die Datei) muss existieren!
    * @method saveAs
    * @param {String} path
    * @param {String} type
    * @param {Boolean} embedFonts
    * @chainable
    *
    * @example
    * Speichert die Präsentation unter dem übergebenen Pfad als 'test.pptx' Datei ab, ohne verwendete Schriftarten abzuspeichern.
    * @example
    *     $presentation.saveAs('C:\\Wunder\\Toller\\Ordner\\test.pptx', 'pptx', false);
    */
    saveAs: function (path, type, embedFonts) {
        PPTNET.saveAs(path, type, embedFonts);
        return this;
    },
    
    
    /**
    * Speichert die PowerPoint-Präsentation unter dem übergebenen Pfad unter dem übergebenen Namen und Typ ab.
    * WICHTIG: Diese Funktion erstellt keine fehlenden Ordner, der angegebene Pfad (bis auf die Datei) muss existieren!
    * @method saveCopyAs
    * @param {String} path
    * @param {String} type
    * @param {Boolean} embedFonts
    * @chainable
    * @example
    * Speichert die Präsentation unter dem übergebenen Pfad als 'test.pptx' Datei ab, ohne verwendete Schriftarten abzuspeichern.
    * @example
    *     $presentation.saveAs('C:\\Wunder\\Toller\\Ordner\\test.pptx', 'pptx', false));
    */
    saveCopyAs: function (path, type, embedFonts) {
        PPTNET.saveCopyAs(path, type, embedFonts);
        return this;
    }
    , name: PPTNET.name
    , path: PPTNET.path
    , slideHeight: PPTNET.slideHeight
    , slideWidth: PPTNET.slideWidth
    , slideMaster: function (index) {
        return { 'shapes': Shapes(PPTNET.slideMasterShapes(index)) };
    }


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
    return new ReportScript(this.selector, this.context);
};


$.extend(ReportScript, { slides: Slides });
$.extend(ReportScript, { shapes: Shapes });
$.extend(ReportScript, { presentation: Presentation });

module.exports = ReportScript;