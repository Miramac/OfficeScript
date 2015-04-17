/**
 *  OfficeScript.PPT.NET.js
 *
 *  Connection C#/JavaScript for all global (scope: "presentation") Functions
 **/
/**
* @class PPTNET
**/
var PPTNET = {
    /**
       *  PRESENTATION-Functions
       * @method saveCopyAs
       * @param {} path
       * @param {} type
       * @param {} embedFonts
       * @return 
       */
    saveCopyAs: function (path, type, embedFonts) {
        embedFonts = embedFonts || false;
        type = type || 'pptx';
        PPT.SaveCopyAs(path, type, embedFonts);
    }
  ,
    /**
      * Description
      * @method saveCopyAs
      * @param {} path
      * @param {} type
      * @param {} embedFonts
      * @return 
      */
    saveAs: function (path, type, embedFonts) {
        embedFonts = embedFonts || false;
        type = type || 'pptx';
        PPT.SaveAs(path, type, embedFonts);
    }
  ,
    /**
      * Description
      * @method save
      * @return 
      */
    save: function () {
        PPT.Save();
    }
  ,
    /**
      * Description
      * @method name
      * @return MemberExpression
      */
    name: function () {
        return PPT.Name;
    }
  ,
    /**
      * Description
      * @method path
      * @return MemberExpression
      */
    path: function () {
        return PPT.Path;
    }
    ,

    /**
      * Description
      * @method firstSlide
      * @return MemberExpression
      */
    firstSlide: function () {
        return PPT.FirstSlide;
    }
    ,

    /**
      * Description
      * @method lastSlide
      * @return MemberExpression
      */
    lastSlide: function () {
        return PPT.LastSlide;
    }
    ,

    /**
      * Description
      * @method activeSlide
      * @return MemberExpression
      */
    activeSlide: function () {
        return PPT.ActiveSlide;
    }
    ,

    /**
      * Description
      * @method slideAtIndex
      * @return MemberExpression
      */
    slideAtIndex: function (index) {
        return PPT.SlideAtIndex(index);
    }
    ,

    /**
      * Description
      * @method path
      * @return MemberExpression
      */
    slideHeight: function () {
        return PPT.SlideHeight;
    }
    ,
    /**
      * Description
      * @method path
      * @return MemberExpression
      */
    slideWidth: function () {
        return PPT.SlideWidth;
    }
    /**
     *  SLIDE-Functions
     **/
  ,
    /**
      * Description
      * @method findSlides
      * @param {} slideIds
      * @return 
      */
    findSlides: function (slideIds) {
        if (slideIds === "*") { //get all slides
            return PPT.FindSlides();
        } else {
            slideIds = slideIds.split(',');
            return PPT.FindSlides(slideIds);
        }
    }
  ,
    /**
      * Description
      * @method addSlide
      * @param {} position
      * @param {} layout
      * @return CallExpression
      */
    addSlide: function (position, layout) {
        return PPT.AddSlide(position, layout);
    }
  ,
    /**
      * Description
      * @method countSlides
      * @return MemberExpression
      */
    countSlides: function () {
        return PPT.CountSlides;
    }
    /**
     *  SHAPE-Functions
     **/
  ,
    /**
      * Description
      * @method findShapes
      * @param {} shapeIds
      * @param {} slides
      * @return 
      */
    findShapes: function (shapeIds, slides) {
        slides = (slides.slides) ? slides.slides : slides;
        slides = (slides.length) ? slides : [];
        if (shapeIds === "*") { //get all slides
            return PPT.FindShapes(slides);
        } else {
            shapeIds = shapeIds.split(',');
            return PPT.FindShapes(shapeIds, slides);
        }
    }
  ,
    /**
      * Description
      * @method findShapesWithRegex
      * @param {} pattern
      * @return CallExpression
      * @example ^Hello World!$ -> Exact ^ = Start, $ = Ende
      * @example (?i)HelloWorld! -> Case insensitive = (?i) zu beginn
      */
    findShapesWithExpression: function (pattern, range) {
        return PPT.FindShapesWithExpression(pattern, range);
    }
  ,
    /**
      * Description
      * @method addTextbox
      * @param {} slide
      * @param {} options
      * @return CallExpression
      */
    addTextbox: function (slide, options) {
        return PPT.AddTextbox(slide, options);
    }
  ,
    /**
      * Description
      * @method addPicture
      * @param {} slide
      * @param {} options
      * @return CallExpression
      */
    addPicture: function (slide, options) {
        return PPT.AddPicture(slide, options);
    }
   ,
    /**
      * Description
      * @method addPicture
      * @param {} slide
      * @param {} options
      * @return CallExpression
      */
    slideMasterShapes: function (index) {
        return PPT.SlideMasterShapes(index)
    }
};

module.exports = PPTNET;