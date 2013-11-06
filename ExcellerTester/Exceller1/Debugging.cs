using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcellerTest1
{
    class Debugging
    {
        public void Run()
        {
            Console.WriteLine(NextColumnName("A"));
            Console.WriteLine(NextColumnName("B"));
            Console.WriteLine(NextColumnName("C"));
            Console.WriteLine(NextColumnName("Z"));
            Console.WriteLine(NextColumnName("AA"));
            Console.WriteLine(NextColumnName("AB"));
            Console.WriteLine(NextColumnName("CE"));
            Console.WriteLine(NextColumnName("DK"));
            Console.WriteLine(NextColumnName("EY"));
            Console.WriteLine(NextColumnName("EZ"));
        }

        private char NextChar(char input)
        {
            return (char)((int)input + 1);
        }

        private string NextColumnName(String columnName)
        {
            if (String.IsNullOrEmpty(columnName)) {
            throw new ArgumentException("columnName");
            }
            if (columnName.Length == 1)
            {
                if ("Z".Equals(columnName, StringComparison.OrdinalIgnoreCase))
                {
                    return "AA";
                }
                else
                {
                    return (NextChar(columnName[0])).ToString();
                }
            }
            else if (columnName.Length == 2)
            {
                if (columnName[1] == 'z' || columnName[1] == 'Z')
                {
                    char first = NextChar(columnName[0]);
                    return first.ToString() + "A";
                }
                else
                {
                    char second = NextChar(columnName[1]);
                    return columnName[0].ToString() + second.ToString();
                }
            }
            else
            {
                throw new ArgumentException("columnName must be 1 or 2 characters long");
            }
        }

    }
}
