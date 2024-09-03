using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExMerge
{
    internal class Utils
    {
        public static string fullToHalf(string s)
        {
            var converted = Regex.Replace(s, "[０-９]", p => ((char)(p.Value[0] - '０' + '0')).ToString());
            converted = Regex.Replace(converted, "[ａ-ｚ]", p => ((char)(p.Value[0] - 'ａ' + 'a')).ToString());
            converted = Regex.Replace(converted, "[Ａ-Ｚ]", p => ((char)(p.Value[0] - 'Ａ' + 'A')).ToString());
            converted = converted.Replace('　', ' ');
            // remove first space
            if (converted.Length > 0 && converted[0] == ' ')
            {
                converted = converted.Substring(1);
            }
            return converted;
        }
    }
}
