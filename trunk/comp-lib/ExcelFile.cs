using System;
using System.IO;

namespace compare_lib
{
    public class ExcelFile
    {
        public string Name;
        public object Object;

        public string ToUri()
        {
            return (new Uri((string)Object)).AbsoluteUri;
        }

        public string ToSize()
        {
            return (new FileInfo((string)Object).Length / 1024).ToString();
        }

    }
}
