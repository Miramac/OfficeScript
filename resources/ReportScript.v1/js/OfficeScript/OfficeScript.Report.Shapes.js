/**
 *  OfficeScript.Report.Shapes.js
 *
 *  Shapes Core: Selector and Basic Shape-Functions
 **/
    var $ = require('./OfficeScript.Core')
    , PPTNET = require('./OfficeScript.PPT.NET')
    , Attributes = require('./OfficeScript.Report.Shapes.Attributes')
    , Utils = require('./OfficeScript.Report.Shapes.Utils')
    , Static = require('./OfficeScript.Report.Shapes.Static')
    , Paragraphs = require('./OfficeScript.Report.Shapes.Paragraphs')
    , Slides = require('./OfficeScript.Report.Slides')
    , _ = require('../lib/lodash')
    ;

/**
* @class $shapes
* @param {} selector
* @param {} context
* @return NewExpression
*/
var Shapes = (function () {
    var Shapes = function (selector, context) {
        if (selector instanceof Shapes) {
            return selector;
        } else {
            return new Shapes.fn.init(selector, context);
        }
    };
    Shapes.fn = Shapes.prototype = {
        constructor: Shapes,
        /**
        * Initialisiert Shape-Objekte auf welche alle weiteren Funktionen angewendet werden können.
        * @method init
        * @param {} selector
        * @param {} context
        * @chainable
        */
        init: function (selector, context) {
            var i, j, tmpShapes;
            this.shapes = [];
            this.slides = [];
            context = context || [];
            if (!selector) {
                return this;
            }
            if (typeof context === 'string') {
                context = Slides(context);
            }
            if (typeof selector === 'string') {
                selector = selector.split(',');
                return Shapes(selector, context);
            }
            if (selector.toString() === 'OfficeScript.ReportScript.Shape') {
                selector = [selector];
            }
            if ($.isArray(selector)) {
                for (i = 0; i < selector.length; i++) {
                    if ((typeof selector[i] === 'string')) {
                        selector.splice(i, 1, PPTNET.findShapes(selector[i].trim(), context));
                    } else if (selector[i] instanceof Shapes) {
                        tmpShapes = [];
                        for (j = 0; j < selector[i].shapes.length; j++) {
                            tmpShapes.push(selector[i].shapes[j]);
                        }
                        selector.splice(i, 1, tmpShapes);
                    }
                }
                selector = _.flatten(selector, true);
                selector = _.compact(selector);
            }

            this.shapes = selector;
            this.slides = context;
            return this;
        },
        
        
        /**
        * Gibt die Anzahl der aktuell ausgewählten PowerPoint-Objekte aus.
        * @method count
        * @chainable
        *
        * @example
        * Lies die Anzahl der aktuell ausgewählten Folien aus und schreibt diese in die Variable 'shapesCount'.
        * @example
        *     var shapesCount = $shapes('selector1, selector2, selector3').count();
        */
        count: function () {
            return this.shapes.length;
        },
        
        
        /**
        * Wendet für jedes übergebene PowerPoint-Objekt die übergebene Funktion an.
        * @method each
        * @param {Function} callback
        * @param {object} args (only for internal use!)
        * @return CallExpression
        */
        each: function (callback, args) {
            return $(this.shapes).each(callback, args);
        },
        
        
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
        attr: function (name, value, parent, targetName) {
            targetName = targetName || 'shapes';
            parent = parent || this;
            return $.attr(name, value, parent, targetName);
        },
        
        
        /**
        * Setzt die aktuell aktiven Shape-Objekte zurück.
        * @method dispose
        * @chainable
        *
        * @example
        * Setzt die Shapes-Objekte zurück. 
        * @example
        *     $shapes('selctor1, selector2').dispose();
        */
        dispose: function () {
            this.each(function () {
                this.Dispose();
            });
            this.shapes = [];
            return this;
        },
        
        
        /**
        * Sucht nach PowerPoint-Objekten in der Präsentation mit dem übergebenen Text und setzt diese als aktive Shapes-Objekte, sodass weiter Funktionen ausgeführt werden können.
        * @method find
        * @param {String|RegExp} needle
        * @param {String} flags
        * @chainable
        *
        * @example
        * Sucht nach allen Objekten die mit dem Text 'Textbox' anfangen
        * @example
        *     $shapes().find('Textbox', 's');
        *
        * @example
        * Sucht nach allen Objekten die mit dem Text 'Textbox' anfangen und schreibt die Anzahl dieser in die Variable 'shapesCount'.
        * @example
        *     var shapesCount = $shapes().find('Textbox', 's').count();
        *
        * @example
        * Sucht nach allen Objekten auf der Folie 'Slide_258', welche mit dem Text 'textbox_c' anfangen und aufhören, wobei nicht auf Groß- und Kleinschreibung geachtet wird, und schreibt die Anzahl dieser in die Variable 'shapesCount'.
        * @example
        *     var shapesCount = $shapes('*', 'Slide_258').find("textbox_c", "ise").count();
        */        
        find: function (needle, flags) {
            var range = (typeof (this.slides.slides) === 'undefined') ? ['*'] : this.slides.slides;
            var pattern = needle.toString();
            if (needle instanceof RegExp) {
                pattern = pattern.substring(1, pattern.length - 1);
            } else {
                //Escape Regexreserved Characters
                pattern = pattern.replace(/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|]/g, "\\$&");
            }
            if (typeof flags !== 'undefined') {
                //  i: Case Insensitive
                if (flags.indexOf('i') > -1) {
                    pattern = '(?i)' + pattern;
                }
                //  s: Starts with
                //  m: Exact Match
                if (flags.indexOf('s') > -1 || flags.indexOf('m') > -1) {
                    pattern = '^' + pattern;
                }
                //  e: Ends with
                //  m: Exact match
                if (flags.indexOf('e') > -1 || flags.indexOf('m') > -1) {
                    pattern = pattern + '$';
                }
            }
            if (needle instanceof RegExp) {
                return $shapes(PPTNET.findShapesWithExpression(pattern, range));
            }
            if (typeof needle === 'string') {
                return $shapes(PPTNET.findShapesWithExpression(pattern, range));
            }
        },
        
        
        /**
        * Description
        * @method findByAttr
        * @chainable
        *
        * @example 
        *     $shapes('*').findByAttr('CTOBJECTDATA.ID', 'Textbox_A');
        */
        findByAttr: function (type, value) {
            var foundShapes = [];
            if (typeof (type) === 'undefined' || typeof (value) === 'undefined') {
                throw Error('findByAttr: Missing Parameters!');
            }
            this.each(function () {
                var shape = $shapes(this);
                if (typeof shape[type] === 'function') {
                    if (shape[type]() === value) {
                        foundShapes.push(shape);
                    }
                }
            });
            return $shapes(foundShapes);
        },
        /**
          * Description
          * @method findByTag
          * @return
          * @example $shapes('*').findByTag('CTOBJECTDATA.ID', 'Textbox_A');
          */
        findByTag: function (name, value) {
            var foundShapes = [];
            if (typeof (name) === 'undefined' || typeof (value) === 'undefined') {
                throw Error('findByTag: Missing Parameters!');
            }
            this.each(function () {
                var shape = $shapes(this);
                if (shape.tag(name) === value) {
                    foundShapes.push(shape);
                }
            });
            return $shapes(foundShapes);
        }

    };
    Shapes.fn.init.prototype = Shapes.fn;
    return Shapes;
}());

//append Shapes.fn Modules
$.extend(Shapes.fn, Attributes);
$.extend(Shapes.fn, Utils);
$.extend(Shapes.fn, { p: Paragraphs, paragraphs: Paragraphs });

//append Shape static functions
$.extend(Shapes, Static);

module.exports = Shapes;