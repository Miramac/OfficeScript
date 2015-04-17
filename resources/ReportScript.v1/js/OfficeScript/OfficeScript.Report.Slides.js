var $ = require('./OfficeScript.Core')
, PPTNET = require('./OfficeScript.PPT.NET')
, Attributes = require('./OfficeScript.Report.Slides.Attributes')
, Utils = require('./OfficeScript.Report.Slides.Utils')
, Static = require('./OfficeScript.Report.Slides.Static')
, _ = require('../lib/lodash')
;
/**
* @class $slide
* @method Slides
* @param {} selector
* @return NewExpression
*/
var Slides = (function () {
    var Slides = function (selector) {
        if (selector instanceof Slides) {
            return selector;
        } else {
            return new Slides.fn.init(selector);
        }
    };
    Slides.fn = Slides.prototype = {
        constructor: Slides,

        /**
        * Initialisiert Slide-Objekte auf welche alle weiteren Funktionen angewendet werden können.
        * @method init
        * @param {} selector
        * @return ThisExpression
        */
        init: function (selector) {
            var i, j, tmpSlides;
            this.slides = [];
            if (!selector) {
                return this;
            }

            if (typeof selector === 'string') {
                //Special Slide Selectors
                if (selector.substr(0, 1) === ':') {
                    selector = selector.split('=');
                    var command = selector[0]
                    ,   value = selector[1] || '';
                    switch (command) {
                        case ':first':
                            return Slides(PPTNET.firstSlide());
                            break;
                        case ':last':
                            return Slides(PPTNET.lastSlide());
                            break;
                        case ':active':
                            return Slides(PPTNET.activeSlide());
                            break;
                        case ':index':
                            return Slides(PPTNET.slideAtIndex(value));
                            break;
                    }
                } else {
                //Normal Slide Selector
                    selector = selector.split(',');
                    return Slides(selector);
                }
            }

            if (selector.toString() === 'OfficeScript.ReportScript.Slide') {
                selector = [selector];
            }
            if ($.isArray(selector)) {

                for (i = 0; i < selector.length; i++) {
                    if ((typeof selector[i] === 'string')) {
                        selector.splice(i, 1, PPTNET.findSlides(selector[i].trim()));
                    } else if (selector[i] instanceof Slides) {
                        tmpSlides = [];
                        for (j = 0; j < selector[i].slides.length; j++) {
                            tmpSlides.push(selector[i].slides[j]);
                        }
                        selector.splice(i, 1, tmpSlides);
                    }
                }
                selector = _.flatten(selector, true);
                selector = _.compact(selector);
                selector = _.uniq(selector); //Test

            }
            this.slides = selector;
            return this;
        },
        
        
        /**
        * Gibt die Anzahl der ausgewählten PowerPoint-Folien aus, oder gibt die Anzahl der gesammten PowerPoint-Folien der Präsentation aus, wenn keine PowerPoint-Folien ausgewählt wurden.
        * @method count
        * @chainable
        * @example
        * Gibt die gesammte Anzahl an Folien aus und schreibt diese in die Variable 'slidesCount'.
        * @example
        *     var slidesCount = $slides.count();
        * @example
        * Gibt die Anzahl an ausgewählten Folien aus (in diesem Fall 2).
        * @example
        *     $slides('selector1 , selector2').count();
        */
        count: function () {
            return this.slides.length;
        },
        
        
        /**
        * Wendet für jede übergebene PowerPoint-Folie die übergebene Funktion an.
        * @method each
        * @param {} callback
        * @param {} args
        * @chainable
        */
        each: function (callback, args) {
            return $(this.slides).each(callback, args);
        },
        
        
        /**
        * Setzt den Wert eines Attributs einer PowerPoint-Folie auf den übergebenen Wert, oder gibt den Wert eines Attributs einer PowerPoint-Folie aus wenn der Parameter 'value' nicht definiert ist.
        * @method attr
        * @param {String} name
        * @param {String} value
        * @param {Object} parent
        * @param {String} targetName
        * @chainable
        * @example
        * Gibt den Wert des Attributs 'Name' der Folie aus und schreibt diesen in die Variable 'attrName'.
        * @example
        *     var attrName = $slides('selector').attr('Name');
        * @example
        * Setzt den Wert des Attributs 'Name' der Folie auf 'testName'
        * @example
        *     $slides('selector').attr('Name', 'testName');
        */
        attr: function (name, value, parent, targetName) {
            targetName = targetName || 'slides';
            parent = parent || this;
            return $.attr(name, value, parent, targetName);
        },
        
        
        /**
        * Setzt die aktuell aktiven Slides-Objekte zurück.
        * @method dispose
        * @chainable
        * @example
        * Setzt die Folien-Objekte zurück. 
        * @example
        *     $slides('selctor1, selector2').dispose();
        */
        dispose: function () {
            this.each(function () {
                this.Dispose();
            });
            this.slides = [];
            return this;
        },


        /**
          * Noch wird kein Attribut nach außen sichtbar gemacht
          * @method dispose
          * @return
          * @example $slides('*').findByAttr('name', 'Title 1');
          */
        findByAttr: function (type, value) {
            var foundSlides = [];
            if (typeof (type) === 'undefined' || typeof (value) === 'undefined') {
                throw Error('findByAttr: Missing Parameters!');
            }
            this.each(function () {
                var slide = $slides(this);
                if (typeof slide[type] === 'function') {
                    if (slide[type]() === value) {
                        foundSlides.push(slide);
                    }
                }
            });
            return $shapes(foundSlides);
        },


        /**
          * Description
          * @method findByTag
          * @return
          * @example $slides('*').findByTag('Tagname', 'Tagvalue');
          */
        findByTag: function (name, value) {
            var foundSlides = [];
            if (typeof (name) === 'undefined' || typeof (value) === 'undefined') {
                throw Error('findByTag: Missing Parameters!');
            }
            this.each(function () {
                var slide = $slides(this);
                if (slide.tag(name) === value) {
                    foundSlides.push(slide);
                }
            });
            return $shapes(foundSlides);
        },


        /**
          * Description
          * @method prev
          * @return
          * @example $slides('Slide_257').prev();
          */
        prev: function () {
            return $slides(':index=' + (this.index() - 1));
        },


        /**
          * Description
          * @method next
          * @return
          * @example $slides('Slide_257').next();
          */
        next: function () {
            return $slides(':index=' + (this.index() + this.count()));
        }
    };
    Slides.fn.init.prototype = Slides.fn;

    return Slides;
}());

//append Slide Modules
$.extend(Slides.fn, Attributes);
$.extend(Slides.fn, Utils);

//append Slide static functions
$.extend(Slides, Static);

module.exports = Slides;
