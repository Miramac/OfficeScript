/**
 *  OfficeScript.Report.Slides.Attributes.js
 *
 **/

    var $ = require('./OfficeScript.Core')
    , PPTNET = require('./OfficeScript.PPT.NET')
    ;

var Attributes = {

    /**
    * Setzt die ID einer PowerPoint-Folie auf den übergebenen Wert oder gibt die ID einer PowerPoint-Folie zurück, wenn der der Parameter 'id' nicht definiert ist.
    * @method id
    * @param {String} id
    * @chainable
    *
    * @example
    * Gibt die ID der Folie aus und schreibt diese in die Variable 'sildeId'.
    * @example
    *     var slideId = $slides('selector').id();
    *
    * @example
    * Setzt die ID der Folie auf 'Slide_1337'.
    * @example
    *     $slides('selector').id('Slide_1337');
    */
    id: function (id) {
        return this.attr('ID', id);
    },
    
    
    /**
    * Setzt die PowerPoint-Folie an die übergebene Position oder gibt die Position einer PowerPoint-Folie aus, wenn der der Parameter 'pos' nicht definiert wird.
    * @method pos
    * @param {Number} pos
    * @chainable
    *
    * @example
    * Gibt die Position der Folie aus und schreibt diese in die Variable 'sildePos'.
    * @example
    *     var slidePos = $slides('selector').pos();
    *
    * @example
    * Schiebt die Folie an die dritte Stelle.
    * @example
    *     $slides('selector').pos(3);
    */
    pos: function (pos) {
        return this.attr('Pos', pos);
    },
    
    
    /**
    * Gibt die Nummer einer PowerPoint-Folie aus.
    * @method number
    * @chainable
    *
    * @example
    * Gibt die Nummer der Folie aus und schreibt diese in die Variable 'sildeNum'.
    * @example
    *     var slideNum = $slides('selector').number();
    */
    number: function () {
        return this.attr('Number');
    },
    
    
    /**
    * Setzt den übergebenen Tag eines PowerPoint-Folie auf den übergebenen Wert oder gibt den Inhalt eines Tags eines PowerPoint-Folie aus, wenn der Parameter 'value' nicht definiert ist.
    * @method tag
    * @param {String} name
    * @param {String} value
    * @chainable
    *
    * @example
    * Liest den Inhalt des Tags der PowerPoint-Folie aus schreibt diesen in die Variable 'slideTag'.
    * @exmaple
    *     var slideTag = $slides('selector').tag('name');
    *
    * @example
    * Schreibt den in 'value' übergebenen Wert in den Tag des PowerPoint-Folie namens 'name'.
    * @example
    *     $shapes('selector').tag('name', 'value');
    */
    tag: function (name, value) {
        if (typeof value !== 'undefined' && value !== null) {
            this.each(function () {
                this.Tag(name, value);
            })
            return this;
        } else {
            if (this.slides[0]) {
                return this.slides[0].Tag(name);
            }
        }
    }
};

module.exports = Attributes;
