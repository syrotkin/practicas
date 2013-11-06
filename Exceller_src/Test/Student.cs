using System;
using System.Collections.Generic;
using System.Text;
using Taramon.Exceller;
using System.Collections;

namespace Test
{
    [DefaultSheet("Sheet1")]
    class Student
    {
        private string _Name;
        [FromCell("B1")]
        public string Name
        {
            get { return _Name; }
            set { _Name = value; }
        }

        private string _Family;
        [FromCell("B2")]
        public string Family
        {
            get { return _Family; }
            set { _Family = value; }
        }

        private ArrayList _Numbers;
        [FromRange("B3","E3")]
        public ArrayList Numbers
        {
            get { return _Numbers; }
            set { _Numbers = value; }
        }
    }
}
