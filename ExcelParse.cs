using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using NPOI.Util;
using NPOI.HPSF;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using Org.BouncyCastle.Asn1.Utilities;
using System.Collections;

namespace xlsxTools
{
    class ExcelParse
    {

        HSSFWorkbook hw;


        // 
        List<ISheet> sheetList;
        Dictionary<string, DataTable> xlsSheet = new Dictionary<string, DataTable>();

        // 表数据
        Dictionary<string, DataTable> xlsDic = new Dictionary<string, DataTable>();


        string SC_path = string.Empty;

        private void ReadXls(FileInfo file)
        {
            // 文件流
            FileStream fs = new FileStream(SC_path + '/' + file.Name, FileMode.Open, FileAccess.Read);
            Console.WriteLine("Current  Export  {0} ", fs.Name);
            // xls 的工作簿
            XSSFWorkbook wb = new XSSFWorkbook(fs);
            int sheetCount = wb.NumberOfSheets;
            ISheet sheet;

            for (int i = 0; i < sheetCount; i++)
            {
                sheet = wb.GetSheetAt(i);
                if (sheet == null)
                {
                    Console.WriteLine(" {0} is  NULL!! ", fs.Name);
                    return;
                }

                if (sheetList.Contains(sheet))
                {
                    Console.WriteLine(" {0} is  NULL!! ", fs.Name);
                    return;
                }
                else
                {
                    sheetList.Add(sheet);
                }

            }
        }

        private void ReadSheet(ISheet sheet)
        {
            DataTable data = new DataTable();
            IRow firstRow = sheet.GetRow(0); // 第一行
            if (firstRow == null)
                return;

            int cellCount = firstRow.LastCellNum; //  第一行一共有多少列

            for (int i = firstRow.FirstCellNum; i < cellCount; i++)
            {
                ICell iCell = firstRow.GetCell(i);
                string cellValue = iCell.StringCellValue;
                if (iCell != null)
                {
                    if (cellValue.Length <= 0 || cellValue == null)
                    {
                        Console.WriteLine(" {0}  cell  is  NULL!! ", sheet.SheetName);
                        return;
                    }
                    DataColumn column = new DataColumn(cellValue);
                    data.Columns.Add(column);
                }
            }

            int rowCount = sheet.LastRowNum;

            // 第一行名字，第二行數據類型  第三行是否重複   第四行注釋   
            // 主表
            for (int i = 1; i < rowCount; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null)
                {
                    Console.WriteLine("{0}  {1} Row is  null ,Please check it ", sheet.SheetName, i);
                    return;
                }

                DataRow dataRow = data.NewRow();
                for (int j = row.FirstCellNum; j < rowCount; j++)
                {
                    if (row.GetCell(j) != null && row.GetCell(j).ToString().Length > 0)
                    {
                        dataRow[j] = row.Cells[j];
                        //Console.WriteLine(" jjjjj   {0}"  , row.Cells[j]);
                    }
                }
                data.Rows.Add(dataRow);
            }
            if (xlsDic.ContainsKey(sheet.SheetName))
            {
                //   Console.WriteLine("The  {0} KEY  Same  ，出现相同的key ", sheet.SheetName);
                return;
            }
            else
            {
                xlsDic.Add(sheet.SheetName, data);
            }
        }



        public Dictionary<string, DataTable> GetXls()
        {
            return xlsDic;
        }



        // 获取xls文件
        public void SetFilePath(string path)
        {
            SC_path = path;
            DirectoryInfo dicInfo = new DirectoryInfo(path);

            foreach (FileInfo file in dicInfo.GetFiles("*.xlsx"))
            {
                ReadXls(file);
            }
        }


    }
}
