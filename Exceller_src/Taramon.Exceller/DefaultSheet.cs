using System;
using System.Collections.Generic;
using System.Text;

namespace Taramon.Exceller
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class DefaultSheet : Attribute
    {
        private string _SheetName;
        public string SheetName
        {
            get { return _SheetName; }
        }

        public DefaultSheet(string sheetName)
        {
            _SheetName = sheetName;
        }
    }
}
