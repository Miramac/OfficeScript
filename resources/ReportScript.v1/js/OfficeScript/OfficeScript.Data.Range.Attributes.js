/**
 *  OfficeScript.Report.Shapes.Attributes.js
 *
 *  Shapes Attributes: Getter/Setter for all Shape-Attributes
 **/
    var $ = require ('./OfficeScript.Core')
    , _ = require('../lib/lodash')
    ;

/**
* @class $shapes
*
**/

var Attributes = {

    /**
    * Setzt den Text eines PowerPoint-Objekts auf den übergebenen Wert oder gib den Text des PowerPoint-Objekts zurück, wenn der Parameter 'text' nicht definiert ist.
    * @method text
    * @param {String} text 
    * @chainable
    * @example
    *     var shapeText = $shapes('selector').text();    // Liest den Text des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeText'.
    * @example
    *     $shapes('selector').text('Fu Bar');    // Setzt den Text des PowerPoint-Objekts auf 'Fu Bar'.
    */
    val: function (val) {
        if (val.length && val.length>0 && !Array.isArray(val[0])) {
            val = [val];
            console.log(val)
        }
        return  this.attr('Value', val);
    },

    /**
    * Setzt den Text eines PowerPoint-Objekts auf den übergebenen Wert oder gib den Text des PowerPoint-Objekts zurück, wenn der Parameter 'text' nicht definiert ist.
    * @method text
    * @param {String} text 
    * @chainable
    * @example
    *     var shapeText = $shapes('selector').text();    // Liest den Text des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeText'.
    * @example
    *     $shapes('selector').text('Fu Bar');    // Setzt den Text des PowerPoint-Objekts auf 'Fu Bar'.
    */
    formular: function (formular) {
        return this.attr('Formular', val);
    }
};

module.exports = Attributes;