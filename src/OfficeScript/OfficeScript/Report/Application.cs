using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = NetOffice.PowerPointApi;

namespace OfficeScript.Report
{
    class PowerPointApplication : IDisposable
    {

        private PowerPoint.Application application;
        private bool disposed;
        private bool closeApplication;

        public PowerPointApplication()
        {
            this.closeApplication = true;
            this.application = new PowerPoint.Application();
            this.application.Visible = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue;
        }

        // Destruktor
        ~PowerPointApplication()
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
                    if (this.closeApplication)
                    {

                        this.application.Quit();
                    }
                    this.application.Dispose();
                    this.application = null;
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
                open = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return this.Open((string)input);
                    }),
                quit = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.Dispose();
                        return null;
                    })
            };
        }

        private object Open(string name)
        {
            return new Presentation(this.application.Presentations.Open(name)).Invoke();
        }
    }
}
