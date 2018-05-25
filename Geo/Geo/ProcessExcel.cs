using System;
using Microsoft.Office.Interop.Excel;

namespace Geo
{
    class ProcessExcel
    {
        protected static Application objExcelApp;//定义Excel Application对象
        private static Workbooks objExcelWorkBooks;//定义Workbook工作簿集合对象
        protected static Workbook objExcelWorkbook;//定义Excel workbook工作簿对象
        private static Worksheet objExcelWorkSheet;//定义Workbook工作表对象
        private static HttpGetHelper httpGetHelper;//定义HttpGetHelper获取经纬度对象

        public static void Process(string originalFileName, int originalColumn, int targetColumn, int rows)
        {
            try
            {
                string workTmp = originalFileName;
                objExcelApp = new Application();
                objExcelWorkBooks = objExcelApp.Workbooks;
                objExcelWorkbook = objExcelWorkBooks.Open(workTmp, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                objExcelWorkSheet = (Worksheet)objExcelWorkbook.Worksheets[1]; 

                httpGetHelper = new HttpGetHelper();

                for (int i = 2; /*objExcelWorkSheet.Cells[i, originalColumn].Text.ToString() != ""*/i < rows + 2; i++)
                {
                    string address = objExcelWorkSheet.Cells[i, originalColumn].Text.ToString();
                    string str = httpGetHelper.GaoDeAnalysis("key=3e0ded4b2852e194c63565d151c2e606&address=" + address);
                    objExcelWorkSheet.Cells[i, targetColumn] = str;
                }

                string targetFileName = originalFileName.Insert(originalFileName.LastIndexOf('.'), "new");
                
                objExcelWorkbook.SaveAs(targetFileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            finally
            {
                objExcelApp.Quit();
            }
        }

        public static void ProcessSchoolCode()
        {
            string workTmp = @"C:\Users\Administrator\Desktop\学习\高性能计算\Geo\Geo\bin\Debug\高校代码.xlsx";
            objExcelApp = new Application();
            objExcelWorkBooks = objExcelApp.Workbooks;
            objExcelWorkbook = objExcelWorkBooks.Open(workTmp, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            objExcelWorkSheet = (Worksheet)objExcelWorkbook.Worksheets[1];

            for (int i = 1; /*objExcelWorkSheet.Cells[i, 1].Text.ToString() != ""*/i<1016; i++)
            {
                string code = objExcelWorkSheet.Cells[i, 1].Text.ToString();
                
                objExcelWorkSheet.Cells[i, 2] = code.Substring(0, 5);
                objExcelWorkSheet.Cells[i, 3] = code.Substring(5, code.Length - 5);
            }

            string name = @"C:\Users\Administrator\Desktop\学习\高性能计算\Geo\Geo\bin\Debug\高校代码new1.xlsx";
            objExcelWorkbook.SaveAs(name, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }
    }
}
