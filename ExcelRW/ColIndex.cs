using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelRW
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    class ColIndexAttribute : Attribute
    {
        public int Index { get; }
        public ColIndexAttribute(int index)
        {
            this.Index = index;
        }
    }
}