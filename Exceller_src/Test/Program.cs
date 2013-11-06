using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

using Taramon.Exceller;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelMapper mapper = new ExcelMapper();
            Student std1 = new Student();

            string excelFileName = Path.Combine(Directory.GetCurrentDirectory(), "Book1.xls");
            mapper.Read(std1, excelFileName);
            
            double sum = 0.0;
            int n = 0;
            foreach (string number in std1.Numbers)
            {
                sum += Convert.ToDouble(number);
                n++;
            }

            Console.WriteLine("Student: {0} {1}\nAverage of Numbers:{2}", std1.Name, std1.Family, sum / n);
            Console.ReadKey();
        }
    }
}
