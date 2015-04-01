using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = NetOffice.PowerPointApi;

namespace OfficeScript.Report
{
    public class Tags
    {
        private dynamic element;
        public Tags(PowerPoint.Presentation presentation)
        {
            this.element = presentation;
        }
        public Tags(PowerPoint.Slide slide)
        {
            this.element = slide;
        }
        public Tags(PowerPoint.Shape shape)
        {
            this.element = shape;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public object Invoke()
        {
            return new
           {
                get = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return this.Get((string)input);
                    }),
                set = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.Set((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                        return this.Invoke();
                    }),
                remove = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.Delete((string)input);
                        return this.Invoke();
                    })
           };
        }

        public string Get(string name)
        {
            return this.element.Tags[name];
        }
        public void Set(Dictionary<string, object> parameters)
        {
            this.Set((string)parameters["name"], (string)parameters["value"]);
        }
        public void Set(string name, object value)
        {
            this.Delete(name); //Used tags need to be deleted befor set new
            this.element.Tags.Add(name, value.ToString());
        }
        public void Delete(string name)
        {
            this.element.Tags.Delete(name);
        }
    }
}
