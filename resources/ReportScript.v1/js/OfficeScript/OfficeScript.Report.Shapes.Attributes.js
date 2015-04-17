/**
 *  OfficeScript.Report.Shapes.Attributes.js
 *
 *  Shapes Attributes: Getter/Setter for all Shape-Attributes
 **/
    var $ = require ('./OfficeScript.Core')
    , PPTNET = require('./OfficeScript.PPT.NET')
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
    *
    * @example
    * Liest den Text des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeText'.
    * @example
    *     var shapeText = $shapes('selector').text();
    *
    * @example
    * Setzt den Text des PowerPoint-Objekts auf 'Fu Bar'.
    * @example
    *     $shapes('selector').text('Fu Bar');
    */
    text: function(text) {
        return  this.attr('Text', text);
    },
    
    
    /**
    * Setzt den Wert des Abstands nach oben eines PowerPoint-Objekts auf den übergebenen Wert oder gibt den Wert des Abstands nach oben des PowerPoint-Objekts zurück, wenn der Parameter 'top' nicht definiert ist.
    * @method top
    * @param {Number} top
    * @chainable
    *
    * @example
    * Liest den Wert des Abstands nach oben des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeTop'.
    * @example
    *     var shapeTop = $shapes('selector').top();
    *
    * @example
    * Setzt den Wert des Abstands nach oben des PowerPoint-Objekts auf 1337.
    * @example
    *     $shapes('selector').top(1337);
    */
    top: function(top) {
        return  this.attr('Top', top);
    }, 
    
    
    /** 
    * Setzt den Wert des Abstands nach links eines PowerPoint-Objekts auf den übergebenen Wert oder gibt den Wert des Abstands nach links des PowerPoint-Objekts zurück, wenn der Parameter 'left' nicht definiert ist.
    * @method left
    * @param {Number} left
    * @chainable
    *
    * @example
    * Liest den Wert des Abstands nach links des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeLeft'.
    * @example
    *     var shapeLeft = $shapes('selector').left();
    *
    * @example
    * Setzt den Wert des Abstands nach links des PowerPoint-Objekts auf 1337.
    * @example
    *     $shapes('selector').left(1337);
    */
    left: function(left) {
        return  this.attr('Left', left);
    },
  
  
    /**
    * Setzt die Höhe eines PowerPoint-Objekts auf den übergebenen Wert oder gibt den Wert der Höhe des PowerPoint-Objekts zurück, wenn der Parameter 'height' nicht definiert ist.
    * @method height
    * @param {Number} height
    * @chainable
    *
    * @example
    * Liest den Wert der Höhe des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeHeight'. 
    * @example
    *     var shapeHeight = $shapes('selector').height();
    *
    * @example
    * Setzt den Wert der Höhe des PowerPoint-Objekts auf 1337.
    * @example
    *     $shapes('selector').height(1337);
    */
    height: function(height) {
        return this.attr('Height', height);
    }, 
    
    
    /**
    * Setzt die Breite eines PowerPoint-Objekts auf den übergebenen Wert oder gibt den Wert der Breite des PowerPoint-Objekts zurück, wenn der Parameter 'width' nicht definiert ist.
    * @method width
    * @param {Number} width
    * @chainable
    *
    * @example
    * Liest den Wert der Breite des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeWidth'.
    * @example
    *     var shapeWidth = $shapes('selector').width();
    *
    * @example
    * Setzt den Wert der Breite des PowerPoint-Objekts auf 1337.
    * @example
    *     $shapes('selector').width(1337);
    */
    width: function(width) {
        return this.attr('Width', width);
    },
    
    
    /**
    * Rotiert ein PowerPoint-Objekt um den übergebenen Wert in Grad nach rechts oder gibt den Wert der Rotation eines PowerPoint-Objekts zurück, wenn der Parameter 'rotation' nicht definbiert ist.
    * @method rotation
    * @param {Number} rotation
    * @chainable
    *
    * @example
    * Liest den Wert der Rotation des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeRotation'.
    * @example
    *     var shapeRotation = $shapes('selector').rotation();
    *
    * @example
    * Dreht das PowerPoint-Objekt um 90 Grad nach rechts.
    * @example
    *     $shapes('selector').rotation(90);
    */
    rotation: function(rotation) {
        return this.attr('Rotation', rotation);
    },
    
    
    /**
    * Befüllt ein PowerPoint-Objekt mit der Farbe mit dem übergebenen Wert oder gibt den Wert der Farbe mit welcher das PowerPoint-Objekt gefüllt ist aus, wenn der Parameter 'fill' nicht definiert ist.
    * @method fill
    * @param {Number} fill
    * @chainable
    *
    * @example
    * Liest den Wert der Farbe des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeColor'.
    * @example
    *     var shapeColor = $shapes('selector').fill();
    *
    * @example
    *  Befüllt das PowerPoint-Objekt mit der Farbe mit dem Wert '#FF9900'. 
    * @example
    *     $shapes('selector').fill('FF9900');
    */
    fill: function(fill) {
        return this.attr('Fill', fill);
    },
    
    
    /**
    * Setzt den Namen eines PowerPoint-Objekts auf den übergebenen Wert oder gibt den Namen des PowerPoint-Objekts zurück, wenn der Parameter 'name' nicht definiert ist.
    * @method name
    * @param {String} name
    * @chainable
    *
    * @example
    * Liest den Namen des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeName'.
    * @example
    *     var shapeName = $shapes('selector').name(); 
    *
    * @example
    * Ändert den Namen des PowerPoint-Objekts in 'Textbox_1337'. 
    * @example
    *     $shapes('selector').name('Textbox_1337');
    */
    name: function(name) {
        return this.attr('Name', name);
    },
 
 
    /**
    * Setzt die ID eines PowerPoint-Objekts auf den übergebenen Wert oder gibt die ID des PowerPoint-Objekts zurück, wenn der Parameter 'id' nicht definiert ist.
    * @method id
    * @param {String} id
    * @chainable
    *
    * @exmaple
    * Liest die ID des PowerPoint-Objekts aus und schreibt diese in die Variable 'shapeID'.
    * @exmaple
    *     var shapeID = $shapes('selector').id();
    *
    * @exmaple
    * Ändert die ID des PowerPoint-Objekts in 'Textbox_1337'. 
    * @example
    *     $shapes('selector').name('Textbox_1337'); 
    */
    id: function(id) {
        return this.attr('ID', id);
    }, 
    
    
    /**
    * Gibt den nächsten sog. 'Parent' (Elternteil, übergeordnetes Objekt) eines PowerPoint-Objekts zurück.
    * @method parent
    * @chainable
    *
    * @example
    * Liest den Parent des PowerPoint-Objekts aus und schreibt diesen in die Variable 'shapeParent'.
    * @example
    *     var shapeParent = $shapes('selector').parent();
    */
    parent: function() {
        return this.attr('Parent');
    },
  
  
    /**
    * Gibt den Inhalt einer Tabelle (als Objekt) zurück.
    * @method table
    * @chainable
    *
    * @example
    * Liest den Inhalt der Tabelle aus und schreibt diesen in die Variable 'shapeTable'.
    * @example
    *     var shapeTable = $shapes('selector').table(); 
    */
    table: function() {
        return this.attr('Table');
    },
  
  
    /**
    * Gibt die ID der Folie/Slide auf der sich das PowerPoint-Objekt befindet zurück.
    * @method slide
    * @chainable
    *
    * @example
    * Ermittelt die ID der Folie des PowerPoint-Objekts und schreibt diese in die Variable 'shapeSlide'.
    * @example    
    *     var shapeSlide = $shapes('selector').slide(); 
    */
    slide: function() {
        return this.attr('Slide');
    }, 
  

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
    tag: function(name, value) {
        if(typeof value !== 'undefined' && value !== null) {
            this.each(function() {
                this.Tag(name, value);
            });
            return this;
        } else {
            if(this.shapes[0]) {
                return this.shapes[0].Tag(name);
            }
        }
    }
    
};

module.exports = Attributes;