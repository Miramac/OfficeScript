using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = NetOffice.PowerPointApi;
using System.Dynamic;

namespace OfficeScript.Report
{
    class Slide
    {
        private PowerPoint.Slide slide;
        private const OfficeScriptType officeScriptType = OfficeScriptType.Slide;
        private bool disposed;

        public Slide(PowerPoint.Slide slide)
        {
            this.slide = slide;
        }

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
                        return new PowerPointTags(this.slide).Invoke();
                    }),
                remove = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        this.Remove();
                        return null;
                    }),
                duplicate = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return this.Duplicate();
                    }),
                shapes = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return this.Shapes();
                    }),
                addTextbox = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        input = (input == null) ? new Dictionary<string,object>() :  input;
                        return this.AddTextbox((input as IDictionary<string, object>).ToDictionary(d => d.Key, d => d.Value));
                    }),
                getType = (Func<object, Task<object>>)(
                    async (input) =>
                    {
                        return officeScriptType;
                    }
                )
            };
        }

        /// <summary>
        /// Init slide Array
        /// </summary>
        /// <returns></returns>
        private object Shapes()
        {
            List<object> shapes = new List<object>();

            foreach (PowerPoint.Shape pptShape in this.slide.Shapes)
            {
                shapes.Add(new Shape(pptShape).Invoke());
            }

            return shapes.ToArray();
        }

        /// <summary>
        /// Deletes the Slide
        /// </summary>
        private void Remove()
        {
            this.slide.Delete();
        }

        /// <summary>
        /// Duplicate Slide, default position is Slide-Index + 1
        /// </summary>
        private object Duplicate()
        {
            return new Slide(this.slide.Duplicate()[1]).Invoke();
        }

        /// <summary>
        /// Not yet Implemented!
        /// </summary>
        private void Sort()
        {
            throw new NotImplementedException("No sorting Algorithm implemented!");
        }

        /// <summary>
        /// AddTextbox and retrun shape object
        /// </summary>
        private object AddTextbox(IDictionary<string, object> parameters)
        {
            object tmpObject;
            float tmpFloat;

            var orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationHorizontal;
            float left = 0;
            float top = 0;
            float height = 100;
            float width = 100;



            //Try to get Shape options: OFFSCRIPT-2
            if (parameters.TryGetValue("left", out tmpObject))
            {
                if (float.TryParse(tmpObject.ToString(), out tmpFloat))
                {
                    left = tmpFloat;
                }
            }
            if (parameters.TryGetValue("top", out tmpObject))
            {
                if (float.TryParse(tmpObject.ToString(), out tmpFloat))
                {
                    top = tmpFloat;
                }
            }
            if (parameters.TryGetValue("height", out tmpObject))
            {
                if (float.TryParse(tmpObject.ToString(), out tmpFloat))
                {
                    height = tmpFloat;
                }
            }
            if (parameters.TryGetValue("width", out tmpObject))
            {
                if (float.TryParse(tmpObject.ToString(), out tmpFloat))
                {
                    width = tmpFloat;
                }
            }

            if (parameters.TryGetValue("texOrientation", out tmpObject))
            {
                switch (tmpObject.ToString().ToLower())
                {
                    case "horizontal":
                        orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationHorizontal;
                        break;
                    case "downward":
                        orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationDownward;
                        break;
                    case "rotatedfareast":
                        orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationHorizontalRotatedFarEast;
                        break;
                    case "upward":
                        orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationUpward;
                        break;
                    case "vertical":
                        orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationVertical;
                        break;
                    case "verticalfareast":
                        orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationVerticalFarEast;
                        break;
                    case "mixed": //what is mixed??
                        orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationMixed;
                        break;
                    default:
                        orientation = NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationHorizontal;
                        break;

                }
            }



            return new Shape(this.slide.Shapes.AddTextbox(orientation, left, top, width, height)).Invoke();
        }


        #region Properties

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

        #endregion Properties
    }
}
