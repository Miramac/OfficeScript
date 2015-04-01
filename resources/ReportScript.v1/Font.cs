using System;
using NetOffice.OfficeApi.Enums;
using Office = NetOffice.OfficeApi;
using System.Drawing;

namespace OfficeScript.ReportScript
{
    class Font : IDisposable  
    {
        private Office.Font2 font;
        private Paragraph paragraph;
        private bool disposed;

        public Font(Office.Font2 font)
        {
            this.font = font;
        }

        public Font(Paragraph paragraph)
        {
            this.paragraph = paragraph;
        }

        /// <summary>
        /// Create the paragraphs object if undefined
        /// </summary>
        /// <returns></returns>
        private bool Init()
        {
            if (this.font == null && this.paragraph != null)
            {
                this.font = this.paragraph.UnderlyingObject.Font;
            }
            return true;
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
            }
           
        }
        // Destruktor
        ~Font() {
            Dispose(false);
        }

        /// <summary>
        /// Get or Set the Bold-Property for this element.
        /// </summary>
        public bool Bold
        {
            get
            {

                Init();
                return (this.font.Bold == MsoTriState.msoTrue ? true : false);
            }
            set
            {

                Init();
                if (value == true)
                {
                    this.font.Bold = MsoTriState.msoTrue;
                }
                else
                {
                    this.font.Bold = MsoTriState.msoFalse;
                }
            }
        }
        /// <summary>
        /// Get or Set the Italic-Property for this element.
        /// </summary>
        public bool Italic
        {
            get
            {

                Init();
                return (this.font.Italic == MsoTriState.msoTrue ? true : false);
            }
            set
            {

                Init();
                if (value == true)
                {
                    this.font.Italic = MsoTriState.msoTrue;
                }
                else
                {
                    this.font.Italic = MsoTriState.msoFalse;
                }
            }
        }
        /// <summary>
        /// Get or Set the Color-Property for this element.
        /// </summary>
        public string Color
        {
            get
            {

                Init();
                string bgr = "#" + this.font.Fill.ForeColor.RGB.ToString("x6");
                return BGRtoRGB(bgr);
            }
            set
            {

                Init();
                this.font.Fill.ForeColor.RGB = ColorTranslator.FromHtml(BGRtoRGB(value)).ToArgb();
            }
        }

        /// <summary>
        /// Get or Set the Size-Property for this element.
        /// </summary>
        public float Size
        {
            get
            {
                Init();
                return this.font.Size;
            }
            set
            {
                Init();
                this.font.Size = value;
            }
        }
        /// <summary>
        /// Get or Set the Name-Property for this element.
        /// </summary>
        public string Name
        {
            get
            {

                Init();
                return this.font.Name;
            }
            set
            {
                Init();
                this.font.Name = value;
            }
        }

        public void Copy(Font src)
        {
            Init();
            this.Bold = src.Bold;
            this.Italic = src.Italic;
            this.Name = src.Name;
            this.Size = src.Size;
        }

        public Office.Font2 UnderlyingObject
        {
            get
            {
                Init();
                return this.font;
            }
        }


        /// <summary>
        /// Helper for Color because .Net treat color as RGB, while Netoffice (Interop aswell) treats color as BGR
        /// </summary>
        private string BGRtoRGB(string value)
        {
            string b = value.Substring(1, 2);
            string g = value.Substring(3, 2);
            string r = value.Substring(5, 2);
            return "#" + r + g + b;
        }

    }
}
