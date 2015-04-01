using System;
using Office = NetOffice.OfficeApi;
using PowerPoint = NetOffice.PowerPointApi;

namespace OfficeScript.ReportScript
{
    class Paragraph
    {
        private PowerPoint.Shape shape;
        private int start;
        private int length;

        public Paragraph(PowerPoint.Shape shape, int start, int length)
        {

            this.shape = shape;
            this.start = start;
            this.length = length;
        }

        /// <summary>
        /// Get or Set the Text-Property for this element.
        /// </summary>
        public string Text
        {
            get
            {
                return this.shape.TextFrame.TextRange.Paragraphs(this.start,this.length).Text.TrimEnd();
            }
            set
            {
                string text = value;
                while (this.shape.TextFrame.TextRange.Paragraphs().Count < this.start-1)
                {
                    this.shape.TextFrame.TextRange.Paragraphs(this.shape.TextFrame.TextRange.Paragraphs().Count).InsertAfter(Environment.NewLine);
                }
                if (this.shape.TextFrame.TextRange.Paragraphs().Count < this.start)
                {
                    text = Environment.NewLine + text;
                }
                this.shape.TextFrame.TextRange.Paragraphs(this.start, this.length).Text = text;
            }
        }

        public int Count
        {
            get
            {
                return this.shape.TextFrame.TextRange.Paragraphs().Count;
            }
        }

        public void Remove()
        {
            this.shape.TextFrame.TextRange.Paragraphs(this.start, this.length).Delete();
        }
        public Font Font
        {
            get
            {
                return new Font(this);
            }
        }
        public Format Format
        {
            get
            {
                return new Format(this);
            }
        }
        public Office.TextRange2 UnderlyingObject
        {
            get
            {
                return this.shape.TextFrame2.TextRange.Paragraphs(this.start, this.length);
            }
        }
    }
}
