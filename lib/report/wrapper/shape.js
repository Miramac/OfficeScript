var _ = require('lodash');


function Shape(edgeShape){

    var shape = {};
    
    /**
    * Setzt den Wert eines Attributs eines PowerPoint-Objekts auf den übergebenen Wert, oder gibt den Wert eines Attributs eines PowerPoint-Objekts aus wenn der Parameter 'value' nicht definiert ist.
    * @method attr
    * @param {String} name
    * @param {String|Number} value
    * @param {Object} parent
    * @param {String} targetName
    * @chainable
    *
    * @example
    * Gibt den Wert des Attributs 'Name' des Objekts aus und schreibt diesen in die Variable 'attrName'.
    * @example
    *     var attrName = $shapes('selector').attr('Name');
    *
    * @example
    * Setzt den Wert des Attributs 'Name' des Objekts auf 'testName'.
    * @example
    *     $shapes('selector').attr('Name', 'testName');
    */
    shape.attr = function(name, value) {
        if(value) {
            return new Shape(edgeShape.attr({name: name, value: value}, true));
        }
        return edgeShape.attr({name: name}, true);
    };
    
    _.assign(shape, require('./shape.attr'));
    
    
    /**
    * Setzt den übergebenen Tag eines PowerPoint-Objekts auf den übergebenen Wert oder gibt den Inhalt des Tags eines PowerPoint-Objekts aus, wenn der Parameter 'value' nicht definiert ist.
    * @method tag
    * @param {String} name
    * @param {String} value
    * @chainable
    *
    * @example
    * Liest den Inhalt des Tags des PowerPoint-Objekts aus schreibt diesen in die Variable 'shapeTag'.
    * @example
    *     var shapeTag = $shapes('selector').tag('name');
    *
    * @example
    * Schreibt den in 'value' übergebenen Wert in den Tag des PowerPoint-Objekts namens 'name'.
    * @example
    *     $shapes('selector').tag('name', 'value');
    *    
    */
    shape.tag = function(name, value) {
        if(typeof value !== 'undefined' && value !== null) {
             edgeShape.tags(null, true).set({name: name, value: value}, true);
            return this;
        }
        return edgeShape.tags(null, true).get(name, true);
    };
    
    return shape;
}

module.exports = Shape;
