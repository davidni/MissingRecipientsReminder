using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace System
{
    internal static class StringExtensions
    {
        public static bool ContainsWholeWord(this string text, string word, bool ignoreCase)
        {
            return Regex.IsMatch(
                text,
                string.Format(CultureInfo.InvariantCulture, @"\b{0}\b", Regex.Escape(word)),
                ignoreCase ? RegexOptions.IgnoreCase : RegexOptions.None);
        }
    }
}
