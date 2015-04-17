/**
 *  OfficeScript.Report.Shapes.Paragraphs.js
 *
 *  Paragraphs Core: Contains all Paragraph-Functions
 **/
    var $ = require('./OfficeScript.Core')
    , PPTNET = require('./OfficeScript.PPT.NET')
    ;

/**
 * Gibt die, anhand der übergebenen Werte definierten, Anzahl an Paragraphen eines PowerPoint-Objekts zurück.
 * @Class $shapes.paragraphs
 * @param {Number} start [-1]    // Nummer des Start-Paragraphs
 * @param {Number} length [-1]    // Anzahl der nach dem Start-Paragraph angezeigten Paragraphen (Start-Paragraph zählt auch dazu)
 * @chainable
 *
 * @example
 * Gibt alle Paragraphen des ausgewählten PowerPoint-Objekts zurück und schreibt diese in die Variable 'shapeParagraphs'.
 * @example
 *     var shapeParagraphs = $shapes('selector').p();
 *
 * @example
 * Gibt den 2., 3. und 4. Paragraphen des ausgewählten PowerPoint-Objekts zurück.
 * @example
 *     $shapes('selector').p(2,3);
 */
var Paragraphs = function (start, length) {
    start = start || -1;
    length = length || -1;
    var self = this
    , paragraphs = []
    , formats = []
    , fonts = []
    , paragraph
    , i
    ;
    //create paragraphs array (like shapes array in Shapes)
    for (i = 0; i < self.shapes.length; i++) {
        paragraph = self.shapes[i].Paragraph(start, length);
        paragraphs.push(paragraph);
        formats.push(paragraph.Format);
        fonts.push(paragraph.Font);
    }
    return {
        //** Properties
        paragraphs: paragraphs
        , formats: formats
        , fonts: fonts
        //** General functions
        ,
      
      
        /**
        * Setzt den Text eines Paragraphen auf den übergebenen Wert oder gib den Text des Paragraphen zurück, wenn der Parameter 'text' nicht definiert ist.
        * @method text
        * @param {String} text
        * @chainable
        *
        * @example
        * Gibt den Text des Paragraphen des PowerPoint-Objekts aus und schreibt diesen in pText.
        * @example
        *     var pText = $shapes('selector').p('selektor').text();
        *
        * @example
        * Setzt den Text des Paragraphen des PowerPoint-Objekts auf 'Fuurbs'.
        * @example
        *     $shapes('selector').p('selektor').text('Fuurbs');
        */
        text: function (text) {
            return self.attr('Text', text, this, 'paragraphs');
        },
        
        
        /**
        * Gibt die gesammte Anzahl der Paragraphen eines PowerPoint-Objekts aus.
        * @method count
        * @chainable
        *
        * @example
        * Gibt die gesammte Anzahl an Paragraphen des PowerPoint-Objekts aus und schreibt diese in die Variable 'pCount'.
        * @example
        *     var pCount = $shapes('selector').p().count();
        */
        count: function () {
            return self.attr('Count', null, this, 'paragraphs');
        },
        
                
        /**
        * Setzt die Ausrichtung eines Paragraphen eines PowerPoint-Objekts auf den übergebenen Wert oder gib den Wert der Ausrichtung eines Paragraphen eines PowerPoint-Objekts zurück, wenn der Parameter 'align' nicht definiert ist.
        * @method align
        * @param {String} align
        * @chainable
        *
        * @example
        * Gibt die Ausrichtung des Paragraphen des PowerPoint-Objekts aus und schreibt diese in die Variable 'pAlign'.
        * @example
        *     var pAlign = $shapes('selector').p('selektor').align();
        *
        * @example
        * Setzt die Ausrichtung des Paragraphen des PowerPoint-Objekts auf 'center.
        * @example
        *     $shapes('selector').p('selektor').align('center');
        */
        align: function (align) {
            return self.attr('Alignment', align, this, 'formats');
        },
        
        
        /**
        * Kopiert die Formatierungseigenschaften(Ausrichtung und Schrifteigenschaften) eines Paragraphen eines PowerPoint-Objekts auf einen übergebenen Paragraphen oder gibt die Formatierung eines Paragraphen eines PowerPoint-Objekts zurück, wenn der Parameter 'src' nicht definiert ist.
        * @method copyFormat
        * @param {$shape} src
        * @chainable
        *
        * @example
        * Liest die Formatierungseigenschaften des Paragraphen aus und schreibt diese als Objekt in die Variable pFormat.
        * @example
        *     var pFormat = $shapes('selector').p('selector').copyFormat();
        *
        * @example
        * Kopiert die Formatierungseigenschaften des Paragraphen ('selector2') auf einen anderen übergebenen Paragraphen ('selector1').
        * @example
        *     $shapes('selector1').p('selector1').copyFormat($shapes('selector2').p('selector2'));
        */
        copyFormat: function (src) {
            // Get Format
            if (!src) return this;
            // Set Format
            var i;
            for (i = 0; i < this.formats.length; i++) {
                this.formats[i].Copy(src.formats[0]);
            }
            for (i = 0; i < this.fonts.length; i++) {
                this.fonts[i].Copy(src.fonts[0]);
            }
            return this;
        },
        
        
        /**
        * Setzt den Text eines Paragraphen eines PowerPoint-Objekts auf fettgedruckt oder gib zurück, ob ein Paragraph eines PowerPoint-Objekts fettgedruckt ist oder nicht, wenn der Parameter 'bold' nicht definiert ist.
        * @method bold
        * @param {Boolean} bold
        * @chainable
        *
        * @example
        * Liest aus ob der Paragraph fettgedruckt ist und schreibt den Status (true oder false) in die Variable 'pBold'.
        * @example
        *     var pBold = $shapes('selector').p('selector').bold();
        *        
        * @example
        * Setzt den ausgewählten Paragraphen auf fettgedruckt.
        * @example
        *     $shapes('selector').p('selector').bold('true');
        */
        bold: function (bold) {
            return self.attr('Bold', bold, this, 'fonts');
        },
        
        
        /**
        * Setzt den Text eines Paragraphen eines PowerPoint-Objekts auf kursiv oder gib zurück, ob ein Paragraph eines PowerPoint-Objekts kursiv ist oder nicht, wenn der Parameter 'italic' nicht definiert ist.
        * @method italic
        * @param {Boolean} italic
        * @chainable
        *
        * @example
        * Liest aus ob der Paragraph kursiv gedruckt ist und schreibt den Status (true oder false) in die Variable 'pItalic'.
        * @example
        *     var pItalic = $shapes('selector').p('selector').italic();
        *
        * @example
        * Setzt den ausgewählten Paragraphen auf kursiv.
        * @example
        *     $shapes('selector').p('selector').italic('true');
        */
        italic: function (italic) {
            return self.attr('Italic', italic, this, 'fonts');
        },
        
        
        /**
        * Setzt die Schriftfarbe eines Paragraphen eines PowerPoint-Objekts auf die Farbe mit dem übergebenen Wert oder gib der Wert der Farbe eines Paragraphen eines PowerPoint-Objekts zurück, wenn der Parameter 'color' nicht definiert ist.
        * @method color
        * @param {String} color
        * @chainable
        *
        * @example
        * Liest den Wert der Schriftfarbe des Paragraphen aus und schreibt den diesen in die Variable 'pColor'.
        * @example
        *     var pColor = $shapes('selector').p('selector').color(); 
        *
        * @example
        * Setzt die Schriftfarbe des Paragraphen auf die Farbe mit dem Wert '#FF9900'.
        * @example
        *     $shapes('selector').p('selector').color('#FF9900'); 
        */
        color: function (color) {
            return self.attr('Color', color, this, 'fonts');
        },
        
        
        /**
        * Setzt die Schriftgröße eines Paragraphen eines PowerPoint-Objekts auf die übergebene Größe oder gib die Schriftgröße eines Paragraphen eines PowerPoint-Objekts zurück, wenn der Parameter 'size' nicht definiert ist.
        * @method size
        * @param {Number} size
        * @chainable
        *
        * @example
        * Liest die Schriftgröße des Paragraphen aus und schreibt den diesen in die Variable 'pSize'.
        * @example
        *     var pSize = $shapes('selector').p('selector').size();
        *
        * @example
        * Setzt die Schriftgröße des Paragraphen auf 40.
        * @example
        *     $shapes('selector').p('selector').size(40);
        */
        size: function (size) {
            return self.attr('Size', size, this, 'fonts');
        },
        
        
        /**
        * Setzt die Schriftart eines Paragraphen eines PowerPoint-Objekts auf die übergebene Schriftart oder gib die Schriftart eines Paragraphen eines PowerPoint-Objekts zurück, wenn der Parameter 'name' nicht definiert ist.
        * @method name
        * @param {String} name
        * @chainable
        *
        * @example
        * Liest die Schriftart des Paragraphen aus und schreibt den diese in die Variable 'pName'.
        * @example
        *     var pName = $shapes('selector').p('selector').name();
        *
        * @example
        * Setzt die Schriftart des Paragraphen auf 'Calibri'.
        * @example
        *     $shapes('selector').p('selector').name('Calibri');
        */
        name: function (name) {
            return self.attr('Name', name, this, 'fonts');
        },
        
        
        /**
        * Löscht den/die ausgewählten Paragraphen.
        * @method remove
        * @chainable
        *
        * @example
        * Löscht den ausgewählten Paragraphen.
        * @example
        *     $shapes('selector').p('selector').remove();
        */
        remove: function () {
            each(function () {
                this.Remove();
            });
        },
    };
    
    /**
    * Wendet für jeden übergebenen Paragraphen eines PowerPoint-Objekts die übergebene Funktion an.
    * @method each
    * @param {Function} callback
    * @param {Object} args    
    */
    function each(callback, args) {
        return $(paragraphs).each(callback, args);
    }
};

module.exports = Paragraphs;