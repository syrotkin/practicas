using System;
using System.Collections.Generic;
using System.Text;

namespace Taramon.Exceller
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class FromRange : Attribute
    {
        private Category _Category = Category.Formatted;
        public Category Category
        {
            get { return _Category; }
            set { _Category = value; }
        }

        private string _StartCellAddress;
        public string StartCellAddress
        {
            get { return _StartCellAddress; }
        }

        private string _EndCellAddress;
        public string EndCellAddress
        {
            get { return _EndCellAddress; }
        }

        public FromRange(string startCellAddress, string endCellAddress)
        {
            _StartCellAddress = startCellAddress;
            _EndCellAddress = endCellAddress;
        }

        public FromRange(string startCellAddress, string endCellAddress, Category category)
        {
            _StartCellAddress = startCellAddress;
            _EndCellAddress = endCellAddress;
            _Category = category;
        }
    }
}
