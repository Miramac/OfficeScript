var _ = require('lodash');


function Shape(shape){
    var $shape = {};
    
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
    $shape.attr = function(name, value) {
        if(value) {
            return new Shape(shape.attr({name: name, value: value}, true));
        }
        return shape.attr({name: name}, true)
    }
    
    _.assign($shape, require('./shape.attr'));
    
    return $shape;
}

module.exports = Shape;
