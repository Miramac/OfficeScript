using NetOffice.OfficeApi.Enums;
using System;
using System.Drawing;
using System.Windows.Forms;
using Vocatus.Office;
using PowerPoint = NetOffice.PowerPointApi;

namespace OfficeScript.ReportScript
{
    class Shape : IDisposable  
    {
        private Font font;
        private PowerPoint.Shape shape;
        private string shapeIdTagName;
        private string slideIdTagName;
        private bool disposed = false;
        public Shape(PowerPoint.Shape shape, string shapeIdTagName, string slideIdTagName) 
        {
            this.shape = shape;
            this.shapeIdTagName = shapeIdTagName;
            this.slideIdTagName = slideIdTagName;
            if (this.shape.HasTextFrame == MsoTriState.msoTrue)
            {
                this.font = new Font(this.shape.TextFrame2.TextRange.Font);
            }
        }
        
        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose() {
            // wird nur beim ersten Aufruf ausgeführt
            if (!disposed) {
                Dispose(true);
                GC.SuppressFinalize(this);
                disposed = true;
            }
        }
        protected virtual void Dispose(bool disposing) {
            if (disposing) {
                this.font.Dispose();
                this.shape.Dispose();
            }
        }
        // Destruktor
        ~Shape() {
            Dispose(false);
        }

        /// <summary>
        /// Remove the Shape
        /// </summary>
        public void Remove()
        {
            this.shape.Delete();
           // Dispose();
        }
        /// <summary>
        /// Copy this Shape to another Slide
        /// </summary>
        /// <param name="slide">Target Slide</param>
        /// <returns>Shape</returns>
        public Shape Copy(Slide slide)
        {
            //Store Clipboard data
            IDataObject clipboardData = Clipboard.GetDataObject();
            this.shape.Copy();
            Shape newShape = new Shape(slide.UnderlyingObject.Shapes.Paste()[1], this.shapeIdTagName, this.slideIdTagName);
            //restore Clipboard data
            try
            {
                Clipboard.SetDataObject(clipboardData);
            }
            catch { }
            return newShape;
        }
        /// <summary>
        /// Duplicate this Shape
        /// </summary>
        /// <returns>Shape</returns>
        public Shape Duplicate()
        {
            return new Shape(this.shape.Duplicate()[1], this.shapeIdTagName, this.slideIdTagName); ;
        }
        /// <summary>
        /// Set the z-ordering of the Shape.
        /// Parameters are: "forward", "backward", "front", "back", "beforetext", "behindtext"
        /// </summary>
        /// <param name="order">Ordering Command</param>
        public void Zindex(string order)
        {
            switch (order.ToLower())
            {
                case "forward":
                    this.shape.ZOrder(MsoZOrderCmd.msoBringForward);
                    break;
                case "backward":
                    this.shape.ZOrder(MsoZOrderCmd.msoSendBackward);
                    break;
                case "front":
                    this.shape.ZOrder(MsoZOrderCmd.msoBringToFront);
                    break;
                case "back":
                    this.shape.ZOrder(MsoZOrderCmd.msoSendBackward);
                    break;
                case "beforetext":
                    this.shape.ZOrder(MsoZOrderCmd.msoBringInFrontOfText);
                    break;
                case "behindtext":
                    this.shape.ZOrder(MsoZOrderCmd.msoSendBehindText);
                    break;
            }
        }

        /// <summary>
        /// Get Tag Value for this element.
        /// </summary>
        public string Tag(string name)
        {
                return this.shape.Tags[name];
        }
        /// <summary>
        /// set Tag Value for this element.
        /// </summary>
        public void Tag(string name, string value)
        {
            PowerPointHelper.SetTag(this.shape, name, value);
        }

        /// <summary>
        /// Get or Set the Text-Property for this element.
        /// </summary>
        public string Text
        {
            get
            {
                if (this.shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    return this.shape.TextFrame.TextRange.Text;
                }
                else
                {
                    return null;
                }
            }
            set
            {
                if(this.shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    this.shape.TextFrame.TextRange.Text = value;
                }
            }
        }
        /// <summary>
        /// Get or Set the Top-Property for this element.
        /// </summary>
        public float Top {
            get
            {
                return this.shape.Top;
            }
            set
            {
                this.shape.Top = value;
            }
        }
        /// <summary>
        /// Get or Set the Left-Property for this element.
        /// </summary>
        public float Left {
            get
            {
                return this.shape.Left;
            }
            set
            {
                this.shape.Left = value;
            }
        }
        /// <summary>
        /// Get or Set the Height-Property for this element.
        /// </summary>
        public float Height {
            get
            {
                return this.shape.Height;
            }
            set
            {
                this.shape.Height = value;
            }
        }
        /// <summary>
        /// Get or Set the Width-Property for this element.
        /// </summary>
        public float Width {
            get
            {
                return this.shape.Width;
            }
            set
            {
                this.shape.Width = value;
            }
        }
        /// <summary>
        /// Get or Set the Rotation-Property for this element.
        /// </summary>
        public float Rotation
        {
            get
            {
                return this.shape.Rotation;
            }
            set
            {
                this.shape.Rotation = value;
            }
        }
        /// <summary>
        /// Get or Set the Fill-Property for this element.
        /// </summary>
        public string Fill
        {
            get
            {
                string bgr = "#" + this.shape.Fill.ForeColor.RGB.ToString("x6");
                return BGRtoRGB(bgr);
            }
            set
            {
                this.shape.Fill.ForeColor.RGB = ColorTranslator.FromHtml(BGRtoRGB(value)).ToArgb();
            }
        }
        /// <summary>
        /// Helper for Fill because .Net treat color as RGB, while Netoffice (Interop aswell) treats color as BGR
        /// </summary>
        private string BGRtoRGB(string value)
        {
            string b = value.Substring(1, 2);
            string g = value.Substring(3, 2);
            string r = value.Substring(5, 2);
            return "#" + r + g + b;
        }

        /// <summary>
        /// Get or Set the Alt-Text for this element.
        /// </summary>
        public string AltText
        {
            get
            {
                return this.shape.AlternativeText;
            }
            set
            {
                this.shape.AlternativeText = value;
            }
        }
        /// <summary>
        /// Get or Set the OfficeScript ID for this element.
        /// </summary>
        public string ID
        {
            get
            {
                return this.shape.Tags[this.shapeIdTagName]; ;
            }
            set
            {
                PowerPointHelper.SetTag(this.shape, this.shapeIdTagName, value);
            }
        }
        /// <summary>
        /// Get or Set the Name-Property for this element.
        /// </summary>
        public string Name
        {
            get
            {
                return this.shape.Name; ;
            }
            set
            {
                this.shape.Name = value;
            }
        }

        public Font Font
        {
            get
            {
                return this.font;
            }
        }

        public Paragraph Paragraph(int start, int length)
        {
            return  new Paragraph(this.shape, start, length);
        }

        public int ParagraphsCount(int start, int length)
        {
            return this.shape.TextFrame.TextRange.Paragraphs().Count;
        }

        public object[] Table
        {
            get
            {
                object[] returnTable = null;
                int rowCount, columnCount;
                if (this.shape.HasTable == MsoTriState.msoTrue)
                {
                    returnTable = new object[shape.Table.Rows.Count];
                    
                    rowCount = 0;
                    foreach (PowerPoint.Row row in shape.Table.Rows)
                    {
                        Shape[] cells = new Shape[shape.Table.Columns.Count];
                        columnCount = 0;
                        foreach (PowerPoint.Cell cell in row.Cells)
                        {
                            cells[columnCount++] = new Shape(cell.Shape, this.shapeIdTagName, this.slideIdTagName);
                        }
                        returnTable[rowCount++] = cells;
                        
                    }
                }
                return returnTable;
            }
        }

        public Slide Slide
        {
            get {
                object parent = this.shape.Parent; 
                while (true)
                {
                    
                    if(parent.GetType().Equals(typeof(PowerPoint.Slide))) {
                        return new Slide((parent as PowerPoint.Slide), this.slideIdTagName);
                    }
                    parent = (parent as PowerPoint.Shape).Parent;
                }
            }
        }
        public object Parent
        {
            get
            {
                if (this.shape.Parent.GetType().Equals(typeof(PowerPoint.Shape)))
                {
                    return new Shape((PowerPoint.Shape)this.shape.Parent, this.shapeIdTagName, this.slideIdTagName);
                }
                else if (this.shape.Parent.GetType().Equals(typeof(PowerPoint.Slide)))
                {
                    throw new NotImplementedException();
                }
                return null;
            }
        }
        public string toString()
        {
            return this.ToString();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        public void ExportAs(string path, string type)
        {
            PowerPoint.Enums.PpShapeFormat ppShapeFormat = PowerPoint.Enums.PpShapeFormat.ppShapeFormatPNG;
            switch (type.ToLower())
            {
                case "png":
                    ppShapeFormat = PowerPoint.Enums.PpShapeFormat.ppShapeFormatPNG;
                    break;
                default:
                    ppShapeFormat = PowerPoint.Enums.PpShapeFormat.ppShapeFormatPNG;
                    break;
            }
            this.shape.Export(path, ppShapeFormat, 722*3, 542*3, PowerPoint.Enums.PpExportMode.ppRelativeToSlide);
        } 

        public PowerPoint.Shape UnderlyingObject
        {
            get 
            {
                return this.shape;
            }
        }
    }
}
