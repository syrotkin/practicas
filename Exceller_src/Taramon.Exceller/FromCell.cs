using System;
using System.Collections.Generic;
using System.Text;

namespace Taramon.Exceller
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class FromCell : Attribute
    {
        private Category _Category = Category.Formatted;
        public Category Category
        {
            get { return _Category; }
            set { _Category = value; }
        }

        private string _CellAddress;
        public string CellAddress
        {
            get { return _CellAddress; }
        }

        public FromCell(string cellAddress)
        {
            _CellAddress = cellAddress;
        }

        public FromCell(string cellAddress, Category category)
        {
            _CellAddress = cellAddress;
            _Category = category;
        }
    }
}
