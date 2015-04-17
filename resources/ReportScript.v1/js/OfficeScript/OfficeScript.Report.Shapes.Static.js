/**
 *  OfficeScript.Report.Shapes.Static.js
 *
 **/

    var $ = require('./OfficeScript.Core')
    , PPTNET = require('./OfficeScript.PPT.NET')
    ;

var Static = {

    /**
    * Fügt einer übergebenen PowerPoint-Folie ein neues PowerPoint-Element des übergebenen Typs mit der übergebenen ID und den übergebenen Optionen hinzu.
    * @method add
    * @param {$slides} slides
    * @param {String} id
    * @param {object} options
    * @chainable
    *
    * @example
    * Fügt auf der Folie eine neue Textbox mit zufälliger ID hinzu.
    * @example
    *     $shapes.add($slides('selector'); 
    *
    * @example  
    * Fügt auf der Folie ein neues Bild ein, welches die ID 'img1' besitzt und aus dem Dateipfad 'src' stammt ($presentation.path() = Speicherort der Präsentation).
    * @example
    *     $shapes.add($slides('selector'),{id:'img1', type:'picture', src:$presentation.path()+ '\\test.gif'});
    */
    add: function (slides, id, options) {
        slides = (typeof slides === 'string') ? $slides(slides) : slides;
        slides = (slides.slides) ? slides.slides : slides;
        slides = (slides.length) ? slides : [slides];
        options = (typeof id === 'object') ? id : (options || {});
        id = (typeof id === 'string') ? id : ((options.id) ? options.id : 'Shape_' + Math.round(Math.random() * 100000000).toString());
        options.type = (options.type) ? options.type.toLowerCase() : 'textbox';
        var i, shapes = [];
        for (i = 0; i < slides.length; i++) {
            if (options.type === 'textbox') {
                shapes.push(PPTNET.addTextbox(slides[i], options));
            } else if (options.type === 'picture') {
                shapes.push(PPTNET.addPicture(slides[i], options));
            } else {
                throw new Error("Missing Shape Type!")
            }
        }
        //In ein Shapes.fn Objekt umwandeln und ID setzen
        return $shapes(shapes).id(id);
    }
};

module.exports = Static;