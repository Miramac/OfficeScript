using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = NetOffice.PowerPointApi;


namespace OfficeScript.Report
{
    class Presentation : IDisposable
    {

        private bool disposed;
        private PowerPoint.Presentation presentation;

        private bool closePresentation = true;

        public Presentation(PowerPoint.Presentation presentation)
        {
            this.presentation = presentation;
        }

        // Destruktor
        ~Presentation()
        {
            Dispose(false);
        }

        #region Dispose

        // Implement IDisposable.
        // Do not make this method virtual.
        // A derived class should not be able to override this method.
        public void Dispose()
        {
            Dispose(true);
            // This object will be cleaned up by the Dispose method.
            // Therefore, you should call GC.SupressFinalize to
            // take this object off the finalization queue
            // and prevent finalization code for this object
            // from executing a second time.
            GC.SuppressFinalize(this);
        }
        // Dispose(bool disposing) executes in two distinct scenarios.
        // If disposing equals true, the method has been called directly
        // or indirectly by a user's code. Managed and unmanaged resources
        // can be disposed.
        // If disposing equals false, the method has been called by the
        // runtime from inside the finalizer and you should not reference
        // other objects. Only unmanaged resources can be disposed.
        protected virtual void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.
            if (!this.disposed)
            {
                // If disposing equals true, dispose all managed
                // and unmanaged resources.
                if (disposing)
                {
                    if (this.closePresentation)
                    {
                        this.presentation.Saved = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue;
                        this.presentation.Close();
                    }
                    this.presentation.Dispose();

                }

                // Note disposing has been done.
                this.disposed = true;

            }
        }
        #endregion Dispose

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public object Invoke()
        {
            return new
            {
                attr = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return Util.Attr(this, (input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value), Invoke);
                    }),
                tags = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return new Tags(this.presentation).Invoke();
                    }),
                save = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.Save();
                        return null;
                    }
                ),
                saveAs = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.SaveAs((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                        return null;
                    }
                ),
                saveAsCopy = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.SaveAsCopy((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                        return null;
                    }
                ),
                close = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.Dispose();
                        return null;
                    }
                ),
                slides = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return this.Slides();
                    }
                )
            };
        }


        #region Save
        private void Save()
        {
            this.presentation.Save();
        }

        public void SaveAs(Dictionary<string, object> parameters)
        {
            string name = (string)(parameters as Dictionary<string, object>)["name"];
            string type = null;
            object tmp;
            if (parameters.TryGetValue("type", out tmp))
            {
                type = (string)tmp;
            }
            this.SaveAs(name, type);
        }


        public void SaveAs(string fileName, string fileType)
        {
            this.SaveAs(fileName, fileType, false);
        }

        public void SaveAsCopy(Dictionary<string, object> parameters)
        {
            string name = (string)(parameters as Dictionary<string, object>)["name"];
            string type = null;
            object tmp;
            if (parameters.TryGetValue("type", out tmp))
            {
                type = (string)tmp;
            }
            this.SaveAs(name, type, true);
        }

        public void SaveAsCopy(string fileName, string fileType)
        {
            this.SaveAs(fileName, fileType, true);
        }

        public void SaveAs(string fileName, string fileType, bool isCopy)
        {
            PowerPoint.Enums.PpSaveAsFileType pptFileType;
            switch (fileType.ToLower())
            {
                case "pdf":
                    pptFileType = PowerPoint.Enums.PpSaveAsFileType.ppSaveAsPDF;
                    break;
                default:
                    pptFileType = PowerPoint.Enums.PpSaveAsFileType.ppSaveAsPresentation;
                    break;
            }
            if (isCopy)
            {
                this.presentation.SaveCopyAs(fileName, pptFileType);
            }
            else
            {
                this.presentation.SaveAs(fileName, pptFileType);
            }
        }
        #endregion save

        /// <summary>
        /// Init slide Array
        /// </summary>
        /// <returns></returns>
        private object Slides()
        {
            List<object> slides = new List<object>();

            foreach (PowerPoint.Slide pptSlide in this.presentation.Slides)
            {
                slides.Add(new Slide(pptSlide).Invoke());
            }

            return slides.ToArray();
        }

        
        #region Properties

        public string Name
        {
            get
            {
                return this.presentation.Name;
            }
        }
        public string Path
        {
            get
            {
                return this.presentation.Path;
            }
        }
        public string FullName
        {
            get
            {
                return this.presentation.FullName;
            }
        }
        #endregion
    }
}
