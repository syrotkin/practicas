using System;
using System.Collections.Generic;
using System.Text;

namespace Taramon.Exceller
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class UseSheet : Attribute
    {
        private string _SheetName;
        public string SheetName
        {
            get { return _SheetName; }
            set { _SheetName = value; }
        }

        public UseSheet()
        {
        }

        public UseSheet(string sheetName)
        {
            _SheetName = sheetName;
        }
    }
}
