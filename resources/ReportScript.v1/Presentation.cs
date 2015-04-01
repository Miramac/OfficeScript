using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.IO;
using NetOffice.PowerPointApi.Enums;
using Vocatus.Office;
using Office = NetOffice.OfficeApi;
using PowerPoint = NetOffice.PowerPointApi;
using System.Text.RegularExpressions;

namespace OfficeScript.ReportScript
{
    class Presentation
    {
        private PowerPoint.Presentation presentation;
        private string slideIdTagName;
        private string shapeIdTagName;
        public Presentation(PowerPoint.Presentation presentation, string shapeIdTagName, string slideIdTagName)
        {
            this.presentation = presentation;
            this.slideIdTagName = slideIdTagName;
            this.shapeIdTagName = shapeIdTagName;
        }

        /// <summary>
        /// Search in PowerPoint.Presentation for PowerPoint.Slides with an specific ID-Tag value.
        /// All results will be converted into an Slide-Object.
        /// </summary>
        /// <param name="slideIdValues">Search the ID-Tag for this Value</param>
        /// <returns>Slide[]: IMPORTANT: This will return Slide-Objects not PowerPoint.Slide!</returns>
        public Slide[] FindSlides(object[] slideIdValues)
        {
            List<Slide> returnSlides = new List<Slide>();
            List<PowerPoint.Slide> slides = PowerPointHelper.FindSlidesByTag(this.presentation, this.slideIdTagName, slideIdValues);
            foreach (PowerPoint.Slide slide in slides)
            {
                returnSlides.Add(new Slide(this.presentation, slide, slideIdTagName));
            }
            return returnSlides.ToArray();
        }
        /// <summary>
        /// Returns all slides
        /// </summary>
        /// <returns></returns>
        public Slide[] FindSlides()
        {
            List<Slide> returnSlides = new List<Slide>();
            foreach (PowerPoint.Slide slide in this.presentation.Slides)
            {
                returnSlides.Add(new Slide(this.presentation, slide, slideIdTagName));
            }
            return returnSlides.ToArray();
        }

        /// <summary>
        /// Search in the PowerPoint.Slide Array, or if the Array is empty in PowerPoint.Presentation, for
        /// PowerPoint.Shape's with an specific ID-Tag value.
        /// All results will be converted into an Shape-Object.
        /// </summary>
        /// <param name="shapeIdValues"></param>
        /// <param name="slides">Array of PowerPoint.Slide's to search, if empty search in PowerPoint.Presentation</param>
        /// <returns>Shape[]: IMPORTANT: This will return Shape-Objects not PowerPoint.Shape!</returns>
        public Shape[] FindShapes(object[] shapeIdValues, object[] slides)
        {
            List<Shape> returnShapes = new List<Shape>();
            List<PowerPoint.Shape> storage = new List<PowerPoint.Shape>();
            List<PowerPoint.Slide> slideRange;
            if (slides.Length > 0)
            {
                slideRange = ToSlideList(slides);
            }
            else
            {
                slideRange = ToSlideList(presentation.Slides);
            }

            foreach (string shapeIdValue in shapeIdValues)
            {
                storage.AddRange(PowerPointHelper.FindShapesByTag(slideRange, this.shapeIdTagName, shapeIdValue));
            }
            foreach(PowerPoint.Shape shape in storage) {
                returnShapes.Add(new Shape(shape, this.shapeIdTagName, this.slideIdTagName));
            }

            return returnShapes.ToArray();
        }
        /// <summary>
        /// Return all shapes in slides range
        /// </summary>
        /// <param name="slides">Array of PowerPoint.Slide's to search, if empty search in PowerPoint.Presentation</param>
        /// <returns>Shape[]: IMPORTANT: This will return Shape-Objects not PowerPoint.Shape!</returns>
        public Shape[] FindShapes(object[] slides)
        {
            List<Shape> returnShapes = new List<Shape>();
            List<PowerPoint.Slide> slideRange;
            if (slides.Length > 0)
            {
                slideRange = ToSlideList(slides);
            }
            else
            {
                slideRange = ToSlideList(presentation.Slides);
            }
            foreach (PowerPoint.Slide slide in slideRange)
            {
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    returnShapes.Add(new Shape(shape, this.shapeIdTagName, this.slideIdTagName));
                }
            }
            return returnShapes.ToArray();
        }
        /// <summary>
        /// Count all Slides in the Presentation
        /// </summary>
        public int CountSlides
        {
            get
            {
                return this.presentation.Slides.Count;
            }
        }

        /// <summary>
        /// Adds an empty textbox on the given slide
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="options"></param>
        /// <returns></returns>
        public Shape AddTextbox(Slide slide, dynamic options) {
            float left = (HasProperty(options, "left")) ? float.Parse(options.left.ToString()) : 0;
            float top = (HasProperty(options, "top")) ? float.Parse(options.top.ToString()) : 0;
            float width = (HasProperty(options, "width")) ? float.Parse(options.width.ToString()) : 100;
            float height = (HasProperty(options, "height")) ? float.Parse(options.height.ToString()) : 100;
            Shape shape = new Shape(slide.UnderlyingObject.Shapes.AddTextbox(Office.Enums.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height), this.shapeIdTagName, this.slideIdTagName);
            shape.UnderlyingObject.TextFrame.TextRange.Text = "NEW TEXTBOX!";
            return shape;
        }

        /// <summary>
        /// Adds an empty textbox on the given slide
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="options"></param>
        /// <returns></returns>
        public Shape AddPicture(Slide slide, dynamic options)
        {
            Debug.WriteLine("AddPicture");
            string path = (HasProperty(options, "src")) ? options.src.ToString().Trim() : "";

            if (!File.Exists(path))
            {
                throw new Exception("Missing file!");
            }

            float left = (HasProperty(options, "left")) ? float.Parse(options.left.ToString()) : 0;
            float top = (HasProperty(options, "top")) ? float.Parse(options.top.ToString()) : 0;
            float width = (HasProperty(options, "width")) ? float.Parse(options.width.ToString()) : 100;
            float height = (HasProperty(options, "height")) ? float.Parse(options.height.ToString()) : 100;
            Shape shape = new Shape(slide.UnderlyingObject.Shapes.AddPicture(path, Office.Enums.MsoTriState.msoFalse, Office.Enums.MsoTriState.msoTrue, left, top), this.shapeIdTagName, this.slideIdTagName);
            return shape;
        }

        /// <summary>
        /// Add an empty Slide to the Presentation
        /// </summary>
        /// <param name="insertAt">Index of insert Position</param>
        /// <param name="layout">PowerPoint Layout for the new Slide</param>
        public Slide AddSlide(int insertAt, string layout)
        {
            PpSlideLayout ppLayout;
            switch (layout)
            {
                case "blank":
                    ppLayout = PpSlideLayout.ppLayoutBlank;
                    break;
                case "object":
                    ppLayout = PpSlideLayout.ppLayoutTextAndObject;
                    break;
                case "title":
                    ppLayout = PpSlideLayout.ppLayoutTitle;
                    break;
                default:
                    ppLayout = PpSlideLayout.ppLayoutText;
                    break;
            }
            return new Slide(this.presentation, this.presentation.Slides.Add(insertAt, ppLayout), this.slideIdTagName);
        }

        /// <summary>
        /// Find Shapes which contain either a certain Text or match an Regular Expression
        /// </summary>
        /// <param name="pattern">Text-Pattern</param>
        /// <param name="slides">Slide-Range to search</param>
        /// <returns>Shape[]: IMPORTANT: This will return Shape-Objects not PowerPoint.Shape!</returns>
        public Shape[] FindShapesWithExpression(string pattern, object[] slides)
        {
            Regex regexp;
            List<Shape> returnShapes = new List<Shape>();
            try
            {
                regexp = new Regex(pattern);
            }
            catch {
                return returnShapes.ToArray();
            }
            //Search on all Slides
            if(slides[0].ToString() == "*") {
                slides = FindSlides();
            }

            foreach (Slide slide in slides)
            {
                foreach (PowerPoint.Shape shape in slide.UnderlyingObject.Shapes)
                {
                    if (shape.HasTextFrame == Office.Enums.MsoTriState.msoTrue)
                    {
                        if (regexp.IsMatch(shape.TextFrame2.TextRange.Text))
                        {
                            returnShapes.Add(new Shape(shape, this.shapeIdTagName, this.slideIdTagName));
                        }
                    }
                    else if (shape.HasTable == Office.Enums.MsoTriState.msoTrue)
                    {
                        foreach (PowerPoint.Row row in shape.Table.Rows)
                        {
                            foreach (PowerPoint.Cell cell in row.Cells)
                            {
                                if (regexp.IsMatch(cell.Shape.TextFrame2.TextRange.Text))
                                {
                                    returnShapes.Add(new Shape(cell.Shape, this.shapeIdTagName, this.slideIdTagName));
                                }
                            }
                        }
                    }
                }
            }

            return returnShapes.ToArray();
        }



        public Shape[] SlideMasterShapes(object index)
        {

            List<Shape> returnShapes = new List<Shape>();
            var design = this.presentation.Designs[index];
            if (design != null)
            {
                foreach (PowerPoint.Shape shape in design.SlideMaster.Shapes)
                {
                    returnShapes.Add(new Shape(shape, this.shapeIdTagName, this.slideIdTagName));
                }
            }
            return returnShapes.ToArray();
        }



        /// <summary>
        /// Save Presentation in a copy, but still using the old name/path
        /// </summary>
        /// <param name="path"></param>
        /// <param name="type"></param>
        /// <param name="embedFonts"></param>
        public void SaveCopyAs(string path, string type, bool embedFonts)
        {
            PpSaveAsFileType fileType;

            switch (type.ToLower())
            {
                case "pdf":
                    fileType = PpSaveAsFileType.ppSaveAsPDF;
                    break;
                case "ppt":
                    fileType = PpSaveAsFileType.ppSaveAsPresentation;
                    break;
                case "pptx":
                    fileType = PpSaveAsFileType.ppSaveAsDefault;
                    break;
                default:
                    fileType = PpSaveAsFileType.ppSaveAsDefault;
                    break;
            }

            this.presentation.SaveCopyAs(
                path
                , fileType
                , ((embedFonts) ? Office.Enums.MsoTriState.msoTrue : Office.Enums.MsoTriState.msoFalse)
            );
        }
        public void SaveCopyAs(string path, string type)
        {
            SaveCopyAs(path, type, false);
        }
        public void SaveCopyAs(string path)
        {
            SaveCopyAs(path, "pptx", false);
        }

        /// <summary>
        /// Save PPT File
        /// </summary>
        /// <param name="path"></param>
        /// <param name="type"></param>
        /// <param name="embedFonts"></param>
        public void SaveAs(string path, string type, bool embedFonts)
        {
            PpSaveAsFileType fileType;

            switch (type.ToLower())
            {
                case "pdf":
                    fileType = PpSaveAsFileType.ppSaveAsPDF;
                    break;
                case "ppt":
                    fileType = PpSaveAsFileType.ppSaveAsPresentation;
                    break;
                case "pptx":
                    fileType = PpSaveAsFileType.ppSaveAsDefault;
                    break;
                default:
                    fileType = PpSaveAsFileType.ppSaveAsDefault;
                    break;
            }

            this.presentation.SaveCopyAs(
                path
                , fileType
                , ((embedFonts) ? Office.Enums.MsoTriState.msoTrue : Office.Enums.MsoTriState.msoFalse)
            );
        }
        public void SaveAs(string path, string type)
        {
            SaveAs(path, type, false);
        }
        public void SaveAs(string path)
        {
            SaveAs(path, "pptx", false);
        }
        //Quick save
        public void Save()
        {
            this.presentation.Save();
        }


        public string Path
        {
            get { return this.presentation.Path; }
        }
        public string Name
        {
            get { return this.presentation.Name; }
        }
        public Slide FirstSlide
        {
            get { return new Slide(this.presentation, this.presentation.Slides[1], slideIdTagName); }
        }
        public Slide LastSlide
        {
            get { return new Slide(this.presentation, this.presentation.Slides[this.presentation.Slides.Count], slideIdTagName); }
        }
        public Slide ActiveSlide
        {
            get {
                PowerPoint.SlideRange active = this.presentation.Application.ActiveWindow.Selection.SlideRange;
                return new Slide(this.presentation, active[1], slideIdTagName);
            }
        }
        public Slide SlideAtIndex(int index)
        {
            Slide erg = null;
            try
            {
                erg = new Slide(this.presentation, this.presentation.Slides[index], slideIdTagName);
            } catch (Exception) { }
            return erg;
        }

        public float SlideHeight
        {
            get { return this.presentation.PageSetup.SlideHeight; }
        }

        public float SlideWidth
        {
            get { return this.presentation.PageSetup.SlideWidth; }
        }

        public object UnderlyingObject
        {
            get { return this.presentation; }
        }


        public void DestroyObject(object obj)
        {
            try
            {
                obj = null;
            }
            catch { }
        }

        /** 
         * private utils
         **/
        #region privateUtils
        private List<PowerPoint.Slide> ToSlideList(PowerPoint.Slides slides)
        {
            List<PowerPoint.Slide> list = new List<PowerPoint.Slide>();
            foreach (PowerPoint.Slide slide in slides)
            {
                list.Add(slide);
            }
            return list;
        }
        private List<PowerPoint.Slide> ToSlideList(object[] slides)
        {
            List<PowerPoint.Slide> list = new List<PowerPoint.Slide>();
            foreach (Slide slide in slides)
            {
                list.Add(slide.UnderlyingObject);
            }
            return list;
        }
        private static bool HasProperty(ExpandoObject expando, string key)
        {
            return ((IDictionary<string, Object>)expando).ContainsKey(key);
        }
        #endregion
    }
}