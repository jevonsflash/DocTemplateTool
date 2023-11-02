using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace Word.Helper
{
    public class CommonHelper
    {

        public static string GetDirectoryPathOrNull(Assembly assembly)
        {
            var location = assembly.Location;
            if (location == null)
            {
                return null;
            }

            var directory = new FileInfo(location).Directory;
            if (directory == null)
            {
                return null;
            }

            return directory.FullName;
        }
        public static T GetRandom<T>(IList<T> a)
        {
            Random rnd = new Random();
            int index = rnd.Next(a.Count);
            return a[index];
        }


        public static IList<T> GetRandoms<T>(IList<T> a, int maxNumber, bool isAmountRandom = true)
        {
            Random rnd = new Random();
            int number = isAmountRandom ? rnd.Next(Math.Min(a.Count, maxNumber)) : maxNumber;
            var tags = new List<T>();
            for (int i = 0; i < number; i++)
            {
                var current = GetRandom(a);
                if (!tags.Contains(current))
                {
                    tags.Add(current);
                }
            }
            return tags;
        }

        public static string GetRandomCellphoneNumber()
        {
            string[] telStarts = "134,135,136,137,138,139,150,151,152,157,158,159,130,131,132,155,156,133,153,180,181,182,183,185,186,176,187,188,189,177,178".Split(',');
            var ran = new Random();
            int n = ran.Next(10, 1000);
            int index = ran.Next(0, telStarts.Length - 1);
            string first = telStarts[index];
            string second = (ran.Next(100, 888) + 10000).ToString().Substring(1);
            string thrid = (ran.Next(1, 9100) + 10000).ToString().Substring(1);
            return first + second + thrid;
        }


        public static string GetRandomCaptchaNumber()
        {
            var ran = new Random();
            var result = (ran.Next(1, 99999) + 100000).ToString();
            return result;
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
