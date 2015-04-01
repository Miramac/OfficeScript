using System;
using NetOffice.OfficeApi.Enums;
using Office = NetOffice.OfficeApi;

namespace OfficeScript.ReportScript
{
    class Format
    {
        private Office.ParagraphFormat2 format;
        private Paragraph paragraph;
private  bool disposed;

        public Format(Office.ParagraphFormat2 format)
        {
            this.format = format;
        }
        public Format(Paragraph paragraph)
        {
            this.paragraph = paragraph;
        }

        private bool Init()
        {
            if (this.format == null && this.paragraph != null)
            {
                this.format = this.paragraph.UnderlyingObject.ParagraphFormat;
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
                // Freigabe verwalteter Objekte
            }
        }
        // Destruktor
        ~Format() {
            Dispose(false);
        }

        /// <summary>
        /// Get or Set the Alignment-Property for this element.
        /// Parameters are: "left", "right", "center"
        /// </summary>
        public string Alignment
        {
            get
            {

                Init();
                switch (this.format.Alignment)
                {
                    case MsoParagraphAlignment.msoAlignLeft:
                        return "left";
                    case MsoParagraphAlignment.msoAlignRight:
                        return "right";
                    case MsoParagraphAlignment.msoAlignCenter:
                        return "center";
                    default:
                        return this.format.Alignment.ToString();
                }
            }
            set
            {

                Init();
                switch (value.ToLower())
                {
                    case "left":
                        this.format.Alignment = MsoParagraphAlignment.msoAlignLeft;
                        break;
                    case "right":
                        this.format.Alignment = MsoParagraphAlignment.msoAlignRight;
                        break;
                    case "center":
                        this.format.Alignment = MsoParagraphAlignment.msoAlignCenter;
                        break;
                }
            }
        }

        /// <summary>
        /// Get or Set the Bullet-Property for this element.
        /// </summary>
        public int Bullet
        {
            get
            {

                Init();
                return (int)this.format.Bullet.Character;
            }
            set
            {

                Init();
                this.format.Bullet.Character = value;
            }
        }

        /// <summary>
        /// Get or Set the Indent-Property for this element.
        /// </summary>
        public int IndentLevel
        {
            get
            {

                Init();
                return this.format.IndentLevel;
            }
            set
            {

                Init();
                this.format.IndentLevel = value;
            }
        }

        public Office.ParagraphFormat2 UnderlyingObject
        {
            get
            {
                Init();
                return this.format;
            }
        }

        //http://codereview/#3WF
        public void Copy(Format src)
        {
            Init();
            Office.ParagraphFormat2 srcFormat = src.UnderlyingObject;

            //Bullets
            this.format.Bullet.Font.Name = srcFormat.Bullet.Font.Name;
            this.format.Bullet.Font.Bold = srcFormat.Bullet.Font.Bold;
            this.format.Bullet.Font.Size = srcFormat.Bullet.Font.Size;
            this.format.Bullet.Font.Fill.ForeColor = srcFormat.Bullet.Font.Fill.ForeColor;
            this.format.Bullet.Character = srcFormat.Bullet.Character;
            this.format.Bullet.RelativeSize = srcFormat.Bullet.RelativeSize;
            this.format.Bullet.Visible = srcFormat.Bullet.Visible;
            //Indent
            this.format.FirstLineIndent = srcFormat.FirstLineIndent;
            this.format.IndentLevel = srcFormat.IndentLevel;
            this.format.LeftIndent = srcFormat.LeftIndent;
            this.format.HangingPunctuation = srcFormat.HangingPunctuation;
            this.format.LineRuleBefore = srcFormat.LineRuleBefore;
            this.format.LineRuleAfter = srcFormat.LineRuleAfter;
            //Spacing
            this.format.SpaceBefore = srcFormat.SpaceBefore;
            this.format.SpaceAfter = srcFormat.SpaceAfter;
            this.format.SpaceWithin = srcFormat.SpaceWithin;
        }
    }
}
