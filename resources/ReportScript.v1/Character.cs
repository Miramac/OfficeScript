using Office = NetOffice.OfficeApi;
using PowerPoint = NetOffice.PowerPointApi;

namespace OfficeScript.ReportScript
{
    class Character
    {
        private Font font;
        private Office.TextRange2 character;

        public Character(PowerPoint.Shape shape, int start, int length)
        {
            this.character = shape.TextFrame2.TextRange.Characters(start, length);
            this.font = new Font(this.character.Font);
        }

        public Font Font
        {
            get { return this.font; }
        }

    }
}
