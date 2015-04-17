var _ = require('lodash');


function Shape(shape){
    var $shape = {};
    
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
