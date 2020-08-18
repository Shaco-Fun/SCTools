using System;
using System.Collections.Generic;
using System.Data;

namespace xlsxTools
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelParse ep = new ExcelParse();
            ep.SetFilePath("G:/xls");

            Dictionary<string , DataTable> xlsData =  ep.GetXls();

            Console.WriteLine(xlsData.Count  );
            Console.WriteLine(xlsData);
            foreach (var item in xlsData)
            {
                Console.WriteLine(item + "/n" );
            }

        }
    }
}
