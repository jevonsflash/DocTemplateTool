using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace DocTemplateTool.Helper
{
    public class CommonHelper
    {
        public static dynamic ToCollections(object o)
        {
            if (o is JObject jo) return jo.ToObject<IDictionary<string, object>>().ToDictionary(k => k.Key, v => ToCollections(v.Value));
            if (o is JArray ja) return ja.ToObject<List<IDictionary<string, object>>>().Select(ToCollections).ToList();
            return o;
        }

        public static bool IsUrl(string toCheck)
        {
            return Regex.IsMatch(toCheck, @"^(http[s]?://)?([\da-z.-]+)\.([a-z.]{2,6})([/\w .-]*)*\/?$");

        }

        public static bool IsBase64(string toCheck)
        {
            if (string.IsNullOrEmpty(toCheck))
            {
                return false;
            }
            toCheck = toCheck.Trim();
            return toCheck.Length % 4 == 0 && Regex.IsMatch(toCheck, @"^[a-zA-Z0-9\+/]*={0,3}$", RegexOptions.None);
        }

    }
}
