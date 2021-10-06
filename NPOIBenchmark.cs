using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Diagnostics;
using System.IO;

namespace GcExcelPerformanceTest
{
    class NPOIBenchmark
    {

        public static void TestSetRangeValues_Double(int rowCount, int columnCount, ref double setTime, ref double getTime, ref double saveTime, ref double usedMem)
        {
            Console.WriteLine();
            Console.WriteLine("NPOI：");
            Console.WriteLine(string.Format("{0} 行 {1} 列で数値（double）を使用する場合のベンチマーク", rowCount, columnCount));

            XSSFWorkbook workbook = new XSSFWorkbook();
            var worksheet = workbook.CreateSheet("poi");

            Random rand = new Random();
            DateTime start = DateTime.Now;

            for (int r = 0; r < rowCount; r++)
            {
                IRow row = worksheet.CreateRow(r);
                for (int c = 0; c < columnCount; c++)
                {
                    row.CreateCell(c).SetCellValue(r + c);
                }
            }
            DateTime end = DateTime.Now;

            setTime = (end - start).TotalSeconds;
            Console.WriteLine(string.Format("数値（double）を設定する：{0:N3} 秒", setTime));

            start = DateTime.Now;
            for (int r = 0; r < rowCount; r++)
            {
                IRow row = worksheet.GetRow(r);
                for (int c = 0; c < columnCount; c++)
                {
                    double d = row.GetCell(c).NumericCellValue;
                }
            }
            end = DateTime.Now;
            getTime = (end - start).TotalSeconds;
            Console.WriteLine(string.Format("数値（double）を取得する：{0:N3} 秒", getTime));

            start = DateTime.Now;
            FileStream fileOut = new FileStream("./files/poi-saved-doubles.xlsx", FileMode.Create, FileAccess.Write);
            workbook.Write(fileOut);
            fileOut.Close();
            workbook.Close();
            end = DateTime.Now;
            saveTime = (end - start).TotalSeconds;
            Console.WriteLine(string.Format("数値（double）を保存する：{0:N3} 秒", saveTime));

        }

        public static void TestSetRangeValues_String(int rowCount, int columnCount, ref double setTime, ref double getTime, ref double saveTime, ref double usedMem)
        {
            Console.WriteLine();
            Console.WriteLine(string.Format("{0} 行 {1} 列で文字列値（string）を使用する場合のベンチマーク", rowCount, columnCount));

            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet worksheet = workbook.CreateSheet("poi");


            Random random = new Random();
            string AlphaNumericString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            DateTime start = DateTime.Now;

            for (int r = 0; r < rowCount; r++)
            {
                IRow row = worksheet.CreateRow(r);
                for (int c = 0; c < columnCount; c++)
                {
                    row.CreateCell(c).SetCellValue(AlphaNumericString[random.Next(25)].ToString());
                }
            }
            DateTime end = DateTime.Now;

            setTime = (end - start).TotalSeconds;
            Console.WriteLine(string.Format("文字列値（string）を設定する：{0:N3} 秒", setTime));

            start = DateTime.Now;
            for (int r = 0; r < rowCount; r++)
            {
                IRow row = worksheet.GetRow(r);
                for (int c = 0; c < columnCount; c++)
                {
                    string s = row.GetCell(c).StringCellValue;
                }
            }
            end = DateTime.Now;

            getTime = (end - start).TotalSeconds;
            Console.WriteLine(string.Format("文字列値（string）を取得する：{0:N3} 秒", getTime));

            start = DateTime.Now;
            FileStream fileOut = new FileStream("./files/poi-saved-string.xlsx", FileMode.Create, FileAccess.Write);
            workbook.Write(fileOut);
            fileOut.Close();
            workbook.Close();
            end = DateTime.Now;
            saveTime = (end - start).TotalSeconds;
            Console.WriteLine(string.Format("文字列値（string）を保存する：{0:N3} 秒", saveTime));

        }

        public static void TestSetRangeValues_Date(int rowCount, int columnCount, ref double setTime, ref double getTime, ref double saveTime, ref double usedMem)
        {
            Console.WriteLine();
            Console.WriteLine(string.Format("{0} 行 {1} 列で日付値（date）を使用する場合のベンチマーク", rowCount, columnCount));

            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet worksheet = workbook.CreateSheet("poi");

            DateTime start = DateTime.Now;

            for (int r = 0; r < rowCount; r++)
            {
                IRow row = worksheet.CreateRow(r);
                for (int c = 0; c < columnCount; c++)
                {
                    row.CreateCell(c).SetCellValue(DateTime.Now);
                }
            }
            DateTime end = DateTime.Now;

            setTime = (end - start).TotalSeconds;
            Console.WriteLine(string.Format("日付値（date）を設定する：{0:N3} 秒", setTime));

            start = DateTime.Now;
            for (int r = 0; r < rowCount; r++)
            {
                IRow row = worksheet.GetRow(r);
                for (int c = 0; c < columnCount; c++)
                {
                    DateTime d = row.GetCell(c).DateCellValue;
                }
            }
            end = DateTime.Now;

            getTime = (end - start).TotalSeconds;
            Console.WriteLine(string.Format("日付値（date）を取得する：{0:N3} 秒", getTime));

            start = DateTime.Now;
            FileStream fileOut = new FileStream("./files/poi-saved-doubles.xlsx", FileMode.Create, FileAccess.Write);
            workbook.Write(fileOut);
            fileOut.Close();
            workbook.Close();
            end = DateTime.Now;
            saveTime = (end - start).TotalSeconds;
            Console.WriteLine(string.Format("日付値（date）を保存する：{0:N3} 秒", saveTime));

        }

        public static void TestSetRangeFormulas(int rowCount, int columnCount, ref double setTime, ref double calcTime, ref double saveTime, ref double usedMem)
        {
            Console.WriteLine();
            Console.WriteLine(string.Format("{0} 行 {1} 列で数式を使用する場合のベンチマーク", rowCount, columnCount));

            XSSFWorkbook workbook = new XSSFWorkbook();
            ISheet worksheet = workbook.CreateSheet("poi");

            Random rand = new Random();


            for (int r = 0; r < rowCount; r++)
            {
                IRow row = worksheet.CreateRow(r);
                for (int c = 0; c < 2; c++)
                {
                    row.CreateCell(c).SetCellValue(r + c);
                }
            }

            DateTime start = DateTime.Now;

            for (int r = 0; r < rowCount; r++)
            {
                IRow row = worksheet.GetRow(r);
                for (int c = 2; c < columnCount + 2; c++)
                {
                    ICell cell = row.CreateCell(c);
                    CellReference reference1 = new CellReference(r, c - 2);
                    CellReference reference2 = new CellReference(r, c - 1);
                    cell.CellFormula = string.Format("SUM({0}, {1})", reference1.FormatAsString(), reference2.FormatAsString());
                }
            }

            DateTime end = DateTime.Now;
            setTime = (end - start).TotalSeconds;
            Console.WriteLine(string.Format("数式を設定する：{0:N3} 秒", setTime));

            start = DateTime.Now;
            workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();
            end = DateTime.Now;

            calcTime = (end - start).TotalSeconds;
            Console.WriteLine(string.Format("数式を計算する：{0:N3} 秒", calcTime));

            start = DateTime.Now;
            FileStream fileOut = new FileStream("./files/poi-saved-formulas.xlsx", FileMode.Create, FileAccess.Write);
            workbook.Write(fileOut);
            fileOut.Close();
            workbook.Close();
            end = DateTime.Now;
            saveTime = (end - start).TotalSeconds;
            Console.WriteLine(string.Format("数式を保存する：{0:N3} 秒", saveTime));

        }

        public static void TestBigExcelFile(int rowCount, int columnCount, ref double openTime, ref double calcTime, ref double saveTime, ref double usedMem)
        {
            Console.WriteLine();
            Console.WriteLine(string.Format("多くの数値、数式、スタイルを含む大きなサイズのExcelファイルを使用する場合のベンチマーク"));

            DateTime start = DateTime.Now;
            XSSFWorkbook workbook = new XSSFWorkbook("./files/test-performance.xlsx");
            DateTime end = DateTime.Now;

            openTime = (end - start).TotalSeconds;
            Console.WriteLine(string.Format("Excelファイルを開く：{0:N3} 秒", openTime));

            start = DateTime.Now;
            workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateAll();
            calcTime = (end - start).TotalSeconds;
            end = DateTime.Now;
            calcTime = (end - start).TotalSeconds;
            Console.WriteLine(string.Format("数式を計算する：{0:N3} 秒", calcTime));

            start = DateTime.Now;
            FileStream fileOut = new FileStream("./files/poi-saved-test-performance.xlsx", FileMode.Create, FileAccess.Write);
            workbook.Write(fileOut);
            fileOut.Close();
            workbook.Close();
            end = DateTime.Now;
            saveTime = (end - start).TotalSeconds;
            Console.WriteLine(string.Format("Excelファイルに保存する：{0:N3} 秒", saveTime));

        }
        public static double GetMemory()
        {
            Process proc = Process.GetCurrentProcess();
            long b = proc.PrivateMemorySize64;
            for (int i = 0; i < 2; i++)
            {
                b /= 1024;
            }
            return b;
        }

    }
}
