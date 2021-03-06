﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;

namespace ExcelRW
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ColIndexAttribute : Attribute
    {
        public int Index { get; }
        public ColIndexAttribute(int index)
        {
            this.Index = index;
        }
    }
    [AttributeUsage(AttributeTargets.Property,AllowMultiple =false,Inherited =true)]
    public class ColTypeAttribute:Attribute
    {
        public ColType ColType { get; }
        public ColTypeAttribute(ColType ct)
        {
            this.ColType = ct;
        }
    }
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public class ColNameAttribute : Attribute
    {
        public string ColName { get; }
        public ColNameAttribute(string name)
        {
            this.ColName = name;
        }
    }
    public enum ColType
    {
        T_STR,
        T_BOOL,
        T_DATE,
        T_NUM,
    }
}