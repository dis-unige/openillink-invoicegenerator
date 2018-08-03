﻿using System.Text;

namespace Openillink.InvoiceGenerator.Helpers
{
    public static class StringExtensions
    {
        public static string FixEncoding(this string input)
        {
            var srcEncoding = Encoding.UTF8;
            var destEncoding = Encoding.UTF8;

            var result = Encoding.Convert(srcEncoding, destEncoding, srcEncoding.GetBytes(input));

            return srcEncoding.GetString(result);
        }
    }
}
