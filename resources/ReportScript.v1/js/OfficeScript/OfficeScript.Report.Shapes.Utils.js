/**
 *  OfficeScript.Report.Shapes.Utils.js
 *
 **/

    var $ = require('./OfficeScript.Core')
    , PPTNET = require('./OfficeScript.PPT.NET')
    ;

var Utils = {

    /**
    * Kopiert ein PowerPoint-Objekt in die übergebene Folie und weist diesem die übergebene ID zu, oder kopiert ein PowerPoint-Objekt an Ort und Stelle mit einer zufälligen ID (dublicate()).
    * @method copy
    * @param {String} slide
    * @param {String} newShapeId
    * @chainable
    *
    * @example
    * Kopiert das Objekt(selector1) auf die übergeben Folie (selector2) und weist die ID 'NewShapeId_2' zu.
    * @example
    *     $shapes('selector1').copy('selector2', 'NewShapeId_2);
    *
    * @example
    * Kopiert das Objekt an Ort und Stelle.
    * @example
    *     $shapes('selector')copy();
    */
    copy: function (slide, newShapeId) {
        if (typeof(slide) === 'undefined') {
            return this.duplicate();
        } else if(typeof(slide) === 'string') {
            slide = $slides(slide);
        }
        slide = (typeof slide.slides !== 'undefined') ? slide.slides[0] : slide;
        newShapeId = newShapeId || 'Shape_' + Math.round(Math.random() * 100000000).toString();
        var i, newShapes = [];
        for (i = 0; i < this.count() ; i++) {
            newShapes.push(this.shapes[i].Copy(slide));
        }
        return $shapes(newShapes).id(newShapeId);
    },
    
    
    /**
    * Kopiert ein PowerPoint-Objekt an Ort und Stelle und weist ihm entweder die übergebene ID oder eine zufällige ID zu, wenn der Parameter 'newShapeId' nicht definiert ist.
    * @method duplicate
    * @param {String} newShapeId
    * @chainable
    *
    * @example
    * Kopiert das Objekt an Ort und Stelle.
    * @example
    *     $shapes('selector')dublicate(); 
    */
    duplicate: function (newShapeId) {
        newShapeId = newShapeId || 'Shape_' + Math.round(Math.random() * 100000000).toString();
        var i, newShapes = [];
        for (i = 0; i < this.count() ; i++) {
            newShapes.push(this.shapes[i].Duplicate());
        }
        return $shapes(newShapes).id(newShapeId);
    },
    
    
    /**
    * Entfernt ein PowerPoint-Objekt.
    * @method remove
    * @chainable
    *
    * @example
    * Entfernt das PowerPoint-Objekt.
    * @example
    *     $shapes('selector').remove();
    */
    remove: function () {
        $(this.shapes).each(function () {
            this.Remove();
        });
    },
    
    
    /**
    * Verschiebt ein PowerPoint-Objekt in der Sicht-Ebene um den übergebenen Wert.
    * @method zindex
    * @param {String} order
    * @chainable
    *
    * @example 
    * Verschiebt das Objekt in der Ebene eins nach vorne.
    * @example 
    *     $shapes('selector').zindex('forward'); 
    */
    zindex: function (order) {
        if (typeof order === 'undefined') return this;
        $(this.shapes).each(function () {
            this.Zindex(order);
        });
    },
    
    
    /**
    * Exportiert ein PowerPoint-Objekt in den übergebenen Pfad als übergebenen Dateientyp.
    * @method exportAs
    * @param {String} path
    * @param {String} type
    * @chainable
    *
    * @example
    * Exportiert das Objekt als PNG in den Präsentationsordner unter einem zufälligen Dateinamen.
    * @example
    *     $shapes('selector').exportAs();
    *
    * @example
    * Exportiert das Objekt als PNG in den übergebenen Pfad unter den Dateinamen 'test.png'.
    * @example
    *     $shapes('selector').exportAs('C:\\Wunder\\Toller\\Ordner\\test.png');
    */
    exportAs: function (path, type) {
        path = path || $presentation.path()+'\\tmp_' + Math.round(Math.random() * 10000).toString() + '.png';
        type = (typeof type !== 'undefined') ? type : 'png';
        $(this.shapes).each(function () {
            this.ExportAs(path, type);
        });
    }
};
module.exports = Utils;
