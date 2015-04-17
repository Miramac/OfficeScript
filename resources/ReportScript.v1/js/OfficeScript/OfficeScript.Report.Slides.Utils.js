/**
 *  OfficeScript.Report.Slides.Utils.js
 *
 **/

    var $ = require('./OfficeScript.Core')
    , PPTNET = require('./OfficeScript.PPT.NET')
    ;

var Utils = {

    /**
    * Löscht die ausgewählte PowerPoint-Folie.
    * @method remove
    * @chainable
    *
    * @example
    * Löscht die Folie.
    * @example
    *     $slides('selector').remove(); 
    */
    remove: function () {
        $(this.slides).each(function () {
            this.Remove();
        });
    },
    
    
    /**
    * Verschiebt alle ausgewählten PowerPoint-Folien an die übergebene Position oder gibt die aktuelle Position von PowerPoint-Folien aus, wenn der Parameter 'postion' nicht definiert ist.
    * @method index
    * @param {Number} position
    * @chainable
    *
    * @example
    * Gibt die aktuelle Position der Folie aus.
    * @example
    *     $slides('selector').index();
    *
    * @example
    * Verschiebt die Folie an die vierte Position.
    * @example
    *     $slides('selector').index(4);
    *
    * @example
    * Verschiebt die beiden Folien an die vierte Position (selector1 an Position 4 und selector2 an Position 5)
    * @example
    *     $slides('selector1, selector2').index(4);
    */
    index: function (position) {
        if (typeof position === 'undefined') {
            return this.attr('Pos');
        } else {
            var maxPosition = PPTNET.countSlides();

            //Falls der neue Index kleiner als der alte ist, müssen die Slides der Reihenfolge nach bearbeitet werden
            if (this.attr('Pos') > position) {
                if (position > maxPosition) position = maxPosition;
                $(this.slides).each(function () {
                    this.Move(position);
                    if (position < maxPosition) position++;
                });
            //Falls der neue Index größer als der alte ist, müssen die Slides in umgekehrter Reihenfolge bearbeitet werden
            } else {
                position = position + this.slides.length - 1;
                if (position > maxPosition) position = maxPosition;
                for (var i = this.slides.length; i > 0; i--) {
                    this.slides[i - 1].Move(position);
                    position = position - 1;
                }
            }            
        }
    },
    
    
    /**
    * Kopiert die ausgewählten PowerPoint-Folien an die übergebene Position und weist den Kopien die übergebene ID zu oder kopiert die die ausgewählten PowerPoint-Folien auf die selbe Position unter zufälliger ID, wenn keine Parameter definiert sind.
    * @method copy
    * @param {Number} position
    * @param {String} id
    * @chainable
    *
    * @example
    * Kopiert die Folie an Ort und Stelle und weist der Kopie eine zufällige ID zu.
    * @example
    *     $slides('selector').copy();
    *
    * @example
    *  Kopiert die Folie an die zehnte  Stelle und weist der Kopie eine zufällige ID zu.
    * @example
    *     $slides('selector').copy(10);
    *
    * @example
    * Kopiert die Folie an die achte  Stelle und weist der Kopie die ID 'Slide_1337' zu.
    * @example
    *     $slides('selector').copy(8, 'Slide_1337');
    *
    * @example
    * Kopiert die Folien an die zweite Stelle.
    * @example
    *     $slides('selector1, selector2').copy(2);
    */
    copy: function (position, id) {
        var slides = [];

        $(this.slides).each(function () {
            var maxPosition = PPTNET.countSlides();
            var temp = this.Copy();
            temp.ID = id || 'Slide_' + Math.round(Math.random() * 100000000).toString();
            
            slides.push(temp);
        });
        slides = $slides(slides);
        if (typeof position !== 'undefined') {
            slides.index(position);
        }
        return slides;
    },

    
    /**
    * Not yet implemented!
    * Sortiert die Folien.
    * @method sort
    * @return 
    */
    sort: function () {
        throw new Error('Not yet implemented!');
    },
    
    
    /**
    * Fügt einer PowerPoint-Folie ein neues PowerPoint-Element mit der übergebenen ID un den übergebenen Optionen hinzu oder fügt einer PowerPoint-Folie eine neue Text-Box mit zufälliger ID hinzu wenn keine Parameter definiert sind.
    * @method addShape
    * @param {String} id
    * @param {String} options
    * @chainable
    *
    * @example
    * Fügt der Folie eine neue Text-Box mit zufälliger ID hinzu.
    * @example
    *     $slides('selector').addShape();
    *
    * @example
    * Fügt der Folie eine neue Text-Box mit der ID 'Shape_1337' hinzu.
    * @example
    *     $slides('selector').addShape('Shape_1337');
    */
    addShape: function (id, options) {
        return $shapes.add(this, id, options);
    }
};

module.exports = Utils;
