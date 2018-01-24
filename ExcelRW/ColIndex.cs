using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;

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
    [AttributeUsage(AttributeTargets.Property,AllowMultiple =false,Inherited =true)]
    class ColTypeAttribute:Attribute
    {
        public ColType ColType { get; }
        public ColTypeAttribute(ColType ct)
        {
            this.ColType = ct;
        }
    }
    enum ColType
    {
        T_STR,
        T_BOOL,
        T_DATE,
        T_NUM,
    }
}