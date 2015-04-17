/**
 *  OfficeScript.Report.Slides.Static.js
 *
 **/

    var $ = require('./OfficeScript.Core')
    , PPTNET = require('./OfficeScript.PPT.NET')
    ;

    
var Static = {
    /**
    * Gibt die gesammte Anzahl an PowerPoint-Folien der PowerPoint-Präsentation aus.
    * @method count
    * @chainable
    *
    * @example
    * Liest die Anzahl der Folien aus und schreibt diese in die Variable 'slidesCount'.
    * @example
    *     var slidesCount = $slides.count();
    */
    count: function () {
        return PPTNET.countSlides();
    },
    
    
    /**
    * Erstellt an der übergebenen Stelle eine neue PowerPoint-Folie mit der übergebenen ID und dem übergebem Layout, wenn alle Parameter definiert sind.
    * Erstellt eine neue leere PowerPoint-Folie mit zufälliger ID an letzter Stelle, wenn keine Parameter definiert sind.
    * @method add
    * @param {Number} position
    * @param {String} id
    * @param {String} layout    // 'blank' 'default' 'title' 'object'
    * @chainable
    *
    * @example
    * Erstellt an dritter Stelle eine neue Folie mit der ID 'Slide_1337' mit dem Layout 'title'.
    * @example
    *     $slides.add(3, 'Slide_1337' , 'title');
    *
    * @example
    * Erstellt an letzter Stelle eine neue Folie mit einer zufälligen ID und dem Layout 'blank' (leere Folie).
    * @example
    *     $slides.add();
    *
    * @example
    * Erstellt an fünfter Stelle eine neue Folie mit einer zufälligen ID und dem Layout 'blank' (leere Folie).
    * @example
    *     $slides.add(5);
    *
    * @example
    * Erstellt an zweiter Stelle eine neue Folie mit einer zufälligen ID und dem Layout 'object'.
    * @example    
    *     $slides.add(2, '', 'object'); 
    */
    add: function (position, id, layout) {
        position = position || 'last'; // PPTNET.countSlides() + 1;
        id = id || 'Slide_' + Math.round(Math.random() * 100000000).toString();
        layout = layout || 'blank'; // alt: 'default', 'title'

        if ($.isNumeric(position)) {
            position = parseInt(position, 10);
        } else {
            switch (position) {
                case 'last':
                    position = this.count() + 1;
                    break;
                case 'first':
                    position = 1;
                    break;
                default:
                    position = this.count() + 1;
            }
        }
        return $slides(PPTNET.addSlide(position, layout)).id(id);
    }
};

module.exports = Static;
