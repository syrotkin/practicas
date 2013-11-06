using System;
using System.Collections.Generic;
using System.Text;

namespace Taramon.Exceller
{
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class ToRange : Attribute
    {
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

        public ToRange(string startCellAddress, string endCellAddress)
        {
            _StartCellAddress = startCellAddress;
            _EndCellAddress = endCellAddress;
        }
    }
}
