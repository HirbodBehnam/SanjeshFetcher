using System;
using System.Collections.Generic;
using System.Text;
using System.Web;

namespace SanjeshFetcher
{
    class Helpers
    {
        /// <summary>
        /// Converts an string of persian number (١٢٣٤) to int
        /// </summary>
        /// <param name="input">Number to convert</param>
        /// <returns></returns>
        public static int ParsePersianNumber(string input)
        {
            return Convert.ToInt32(ToEnglishNumber(HttpUtility.HtmlDecode(input).Trim()));
        }
        /// <summary>
        /// Converts numbers like ١،٢،٣،٤ to 1,2,3,4
        /// The other characters are untouched
        /// https://stackoverflow.com/a/30733198/4213397
        /// </summary>
        /// <param name="input">The string to convert</param>
        /// <returns>English string</returns>
        public static string ToEnglishNumber(string input)
        {
            StringBuilder englishNumbers = new StringBuilder(input.Length);
            for (int i = 0; i < input.Length; i++)
            {
                if (char.IsDigit(input[i]))
                    englishNumbers.Append(char.GetNumericValue(input, i));
                else
                    englishNumbers.Append(input[i].ToString());
            }
            return englishNumbers.ToString();
        }
        /// <summary>
        /// Creates an array of ints in a specific range
        /// 1,6 -> [1,2,3,4,5,6]
        /// TODO: should be a better way?
        /// </summary>
        /// <param name="start">The start index</param>
        /// <param name="end">The last index</param>
        /// <returns></returns>
        public static int[] Range(int start, int end)
        {
            var res = new int[end - start + 1];
            for (; start <= end; start++)
                res[end - start] = start;
            return res;
        }
        /// <summary>
        /// Appends two or more arrays together
        /// </summary>
        /// <param name="a">Arrays to join</param>
        /// <returns>Joined array</returns>
        public static int[] AppendArrays(params int[][] a)
        {
            var res = new List<int>();
            foreach (var array in a)
                res.AddRange(array);
            return res.ToArray();
        }
    }
}
