using System;
using Vocatus.Office;
using PowerPoint = NetOffice.PowerPointApi;

namespace OfficeScript.ReportScript
{
    class Slide : IDisposable
    {
        private PowerPoint.Presentation presentation;
        private PowerPoint.Slide slide;
        private string slideIdTagName;
        private bool disposed;

        public Slide(PowerPoint.Presentation presentation, PowerPoint.Slide ppSlide, string slideIdTagName)
        {
            this.presentation = presentation;
            this.slide = ppSlide;
            this.slideIdTagName = slideIdTagName;
        }

        public Slide(PowerPoint.Slide ppSlide, string slideIdTagName)
        {
            this.slide = ppSlide;
            this.slideIdTagName = slideIdTagName;
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
            if (disposing)
            {
                this.slide.Dispose();
            }
        }
        // Destruktor
        ~Slide() {
            Dispose(false);
        }

        /// <summary>
        /// Set the Slide position
        /// </summary>
        /// <param name="position">Index of new Position</param>
        public void Move(int position) {
            this.slide.MoveTo(position);
        }
        /// <summary>
        /// Deletes the Slide
        /// </summary>
        public void Remove() {
            this.slide.Delete();
            Dispose();
        }
        /// <summary>
        /// Copy Slide, default position is Slide-Index + 1
        /// </summary>
        public Slide Copy() {
            return new Slide(this.slide.Duplicate()[1], this.slideIdTagName);
        }
        /// <summary>
        /// Not yet Implemented! (Fabi: worin besteht der Unterschied zu Move()?)
        /// </summary>
        public void Sort()
        {
            throw new NotImplementedException("No sorting Algorithm implemented!");
        }

        /// <summary>
        /// Get Tag Value for this element.
        /// </summary>
        public string Tag(string name)
        {
            return this.slide.Tags[name];
        }
        /// <summary>
        /// Set Tag Value for this element.
        /// </summary>
        public void Tag(string name, string value)
        {
            PowerPointHelper.SetTag(this.slide, name, value);
        }

        public int Pos
        {
            get
            {
                return this.slide.SlideIndex;
            }
            set
            {
                this.slide.MoveTo(value);
            }
        }
        public int Number
        {
            get
            {
                return this.slide.SlideNumber;
            }
        }

        public string Name
        {
            get
            {
                return this.slide.Name;
            }
            set
            {
                this.slide.Name = value;
            }
        }

        public string ID
        {
            get
            {
                return this.slide.Tags[this.slideIdTagName]; ;
            }
            set
            {
                PowerPointHelper.SetTag(this.slide, this.slideIdTagName, value);
            }
        }
        public string toString()
        {
            return this.ToString();
        }
        public PowerPoint.Slide UnderlyingObject
        {
            get
            {
                return this.slide;
            }
        }
    }
}
