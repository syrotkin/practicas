using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Taramon.Exceller;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ExcellerTest1
{
    public class ExcellerTesterApp
    {
       public  ExcelManager ExcelManager { get; set; }

        public void Run()
        {
            String fileName = "myfile.xls";
            

            String filePath = Path.GetFullPath(fileName);

                Console.WriteLine(string.Format("current dir: {0}", Directory.GetCurrentDirectory()));
                ExcelManager.Open(filePath);

                ExcelManager.ActivateSheet("Sheet2");
                //TryReadAddress(excelManager, "A2");

                ReadData();


                //excelManager.SetRangeValues("A4", "C5", new int[] {1, 2, 3});
                //CreateNewSheet(excelManager, "MySheet");
                SaveAs(filePath, "myfile2.xls");
            Console.WriteLine("done");
        }

        private void ReadData()
        {
            // 1. Read the title, store its values
            string columnName = "A";
            List<string> header = ReadHeader(ref columnName);
            int numColumns = header.Count;
            PrintList(header);

            // 2. go down the first column, read the names of specialists, store in a dict; <string, List<List<string>>> 
            // Go to the right, while there are values. Always read formatted values. If the value is empty, break
            Dictionary<string, List<List<string>>> doc2Rows = new Dictionary<string, List<List<string>>>();
            int rowNumber = 2;
            while (rowNumber < 64000)
            {
                string doc = ExcelManager.GetFormattedValue("A" + rowNumber).ToString();
                if (string.IsNullOrEmpty(doc))
                {
                    break;
                }
                if (!doc2Rows.ContainsKey(doc))
                {
                    doc2Rows[doc] = new List<List<string>>();
                }
                List<string> line = ReadLine(rowNumber, numColumns);
                doc2Rows[doc].Add(line);
                ++rowNumber;
            }

            // 3. go through the dict. Foreach key create a sheet, copy the title there, copy all the lists pertaining to this specialist
            foreach (string key in doc2Rows.Keys)
            {
                Worksheet newSheet = ExcelManager.AddWorksheet(key);
                ExcelManager.ActivateSheet(newSheet.Name);
                ExcelManager.SetRangeValues("A1", columnName + "1", header);

                Console.WriteLine("filling in from A1 to {0}", columnName + "1");
                rowNumber = 2;
                foreach (List<string> line in doc2Rows[key])
                {
                    ExcelManager.SetRangeValues("A" + rowNumber, columnName + rowNumber, line);
                    ++rowNumber;
                }
                Console.WriteLine("last one: sheet: {0}, cell: {1}", key, columnName + rowNumber);
            }
        }

        private List<string> ReadLine(int rowNumber, int numColumns)
        {
            List<string> line = new List<string>();
            int counter = 0;
            string columnName = "A";
            while (counter < numColumns)
            {
                if (!AddCellValueToList(line, columnName, rowNumber))
                {
                    break;
                }
                columnName = NextColumnName(columnName);
                ++counter;
            }
            return line;
        }

        private List<string> ReadHeader(ref string columnName)
        {
            List<string> header = new List<string>();
            int counter = 0; // just as a safety check, not to run forever
            //string columnName = "A";
            int rowNumber = 1;

            while (++counter < 64000)
            {
                if (!AddCellValueToList(header, columnName, rowNumber))
                {
                    break;
                }
                columnName = NextColumnName(columnName);
            }

            return header;
        }

        private bool AddCellValueToList(List<string> list, string columnName, int rowNumber)
        {
            object value = ExcelManager.GetFormattedValue(columnName + rowNumber.ToString());
            if (String.IsNullOrEmpty(value.ToString()))
            {
                return false;
            }
            list.Add(value.ToString());
            return true;
        }

        private void PrintList<T>(List<T> list)
        {
            foreach (T item in list)
            {
                Console.Write("{0} ", item);
            }
            Console.WriteLine();
        }

        private void TryReadAddress(string address)
        {
            double? a1NumericValue = null;
            String a1FormattedValue = null; 
            bool isError = false;
            try
            {
                a1NumericValue = ExcelManager.GetNumericValue(address);
            }
            catch (InvalidCastException)  // If there is no numeric value, exception is thrown.
            {
                Console.WriteLine("CastException. Value at address: {0} is not numeric", address);
                isError = true;
            }

            if (!isError)
            {
                if (a1NumericValue.HasValue)
                {
                    Console.WriteLine("Address: {0}, numeric value: {1}", address, a1NumericValue);
                }
            }
            else
            {
                a1FormattedValue = ExcelManager.GetFormattedValue(address).ToString();
                Console.WriteLine("Address: {0}, formatted value: {1}", address, a1FormattedValue);
            }
        }

        private void SaveAs(String filePath, String newFileName)
        {
            String newFilePath = String.Format(@"{0}{1}{2}",
                    filePath.Substring(0, filePath.LastIndexOf(Path.DirectorySeparatorChar)),
                    Path.DirectorySeparatorChar,
                    newFileName);
            Console.WriteLine(String.Format("new file path: {0}", newFilePath));
            ExcelManager.SaveAs(newFilePath);
        }

        private char NextChar(char input)
        {
            return (char)((int)input + 1);
        }

        private string NextColumnName(String columnName)
        {
            if (String.IsNullOrEmpty(columnName))
            {
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
