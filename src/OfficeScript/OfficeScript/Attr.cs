using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeScript
{
    public static class Util
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="parameters"></param>
        /// <returns></returns>
        public static object Attr(object thisObject, Dictionary<string, object> parameters, Func<object> Invoke)
        {
            string name = (string)(parameters as Dictionary<string, object>)["name"];
            object value = null;
            object tmp;
            if (parameters.TryGetValue("value", out tmp))
            {
                value = tmp;
            }

            return Attr(thisObject, name, value, Invoke);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public static object Attr(object thisObject, string name, object value, Func<object> Invoke)
        {
            if (value != null)
            {
                thisObject.GetType().GetProperty(name).SetValue(thisObject, value, null);
                return Invoke();
            }
            else
            {
                return thisObject.GetType().GetProperty(name).GetValue(thisObject, null);
            }
        }
    }
}
