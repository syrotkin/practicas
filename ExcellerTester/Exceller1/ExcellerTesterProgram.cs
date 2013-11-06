using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Taramon.Exceller;

namespace ExcellerTest1
{
    class ExcellerTesterProgram
    {
        static void Main(string[] args)
        {
            //Debugging d = new Debugging();
            //d.Run();
            using (ExcelManager excelManager = new ExcelManager())
            {
                ExcellerTesterApp app = new ExcellerTesterApp();
                app.ExcelManager = excelManager;
                app.Run();
            }

            Console.ReadLine();
        }
    }
}
