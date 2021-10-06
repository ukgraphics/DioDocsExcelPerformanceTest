using GrapeCity.Documents.Excel;
using System;
using System.Diagnostics;

namespace GcExcelPerformanceTest
{
    public class GcExcelBenchmark
    {

        public static void TestSetRangeValues_Double(int rowCount, int columnCount, ref double setTime, ref double getTime, ref double saveTime, ref double usedMem)
        {

            Console.WriteLine("DioDocs for Excel：");
            Console.WriteLine(string.Format("{0} 行 {1} 列で数値（double）を使用する場合のベンチマーク", rowCount, columnCount));

            IWorkbook workbook = new Workbook();
            IWorksheet worksheet = workbook.Worksheets[0];

            DateTime start = DateTime.Now;

            double[,] values = new double[rowCount, columnCount];

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    values[i, j] = i + j;
                }
            }

            worksheet.Range[0, 0, rowCount, columnCount].Value = values;
            DateTime end = DateTime.Now;

            setTime = (end - start).TotalSeconds;
            Console.WriteLine(string.Format("数値（double）を設定する：{0:N3} 秒", setTime));

            start = DateTime.Now;
            object tmpValues = worksheet.Range[0, 0, rowCount, columnCount].Value;
            end = DateTime.Now;

            getTime = (end - start).TotalSeconds; ;
            Console.WriteLine(string.Format("数値（double）を取得する：{0:N3} 秒", getTime));

            start = DateTime.Now;
            workbook.Save("./files/gcexcel-saved-doubles.xlsx");
            end = DateTime.Now;
            saveTime = (end - start).TotalSeconds; ;
            Console.WriteLine(string.Format("数値（double）を保存する：{0:N3} 秒", saveTime));

        }


        public static void TestSetRangeValues_String(int rowCount, int columnCount, ref double setTime, ref double getTime, ref double saveTime, ref double usedMem)
        {

            Console.WriteLine();
            Console.WriteLine(string.Format("{0} 行 {1} 列で文字列値（string）を使用する場合のベンチマーク", rowCount, columnCount));

            IWorkbook workbook = new Workbook();
            IWorksheet worksheet = workbook.Worksheets[0];

            Random random = new Random();
            string AlphaNumericString = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            DateTime start = DateTime.Now;

            string[,] values = new string[rowCount, columnCount];

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    values[i, j] = AlphaNumericString[random.Next(25)].ToString();
                }
            }

            worksheet.Range[0, 0, rowCount, columnCount].Value = values;
            DateTime end = DateTime.Now;

            setTime = (end - start).TotalSeconds; ;
            Console.WriteLine(string.Format("文字列値（string）を設定する：{0:N3} 秒", setTime));

            start = DateTime.Now;
            object tmpValues = worksheet.Range[0, 0, rowCount, columnCount].Value;
            end = DateTime.Now;

            getTime = (end - start).TotalSeconds; ;
            Console.WriteLine(string.Format("文字列値（string）を取得する：{0:N3} 秒", getTime));

            start = DateTime.Now;
            workbook.Save("./files/gcexcel-saved-string.xlsx");
            end = DateTime.Now;
            saveTime = (end - start).TotalSeconds; ;
            Console.WriteLine(string.Format("文字列値（string）を保存する：{0:N3} 秒", saveTime));

        }

        public static void TestSetRangeValues_Date(int rowCount, int columnCount, ref double setTime, ref double getTime, ref double saveTime, ref double usedMem)
        {

            Console.WriteLine();
            Console.WriteLine(string.Format("{0} 行 {1} 列で日付値（date）を使用する場合のベンチマーク", rowCount, columnCount));

            IWorkbook workbook = new Workbook();
            IWorksheet worksheet = workbook.Worksheets[0];

            DateTime start = DateTime.Now;

            DateTime[,] values = new DateTime[rowCount, columnCount];

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < columnCount; j++)
                {
                    values[i, j] = DateTime.Now;
                }
            }
            worksheet.Range[0, 0, rowCount, columnCount].Value = values;
            DateTime end = DateTime.Now;

            setTime = (end - start).TotalSeconds; ;
            Console.WriteLine(string.Format("日付値（date）を設定する：{0:N3} 秒", setTime));

            start = DateTime.Now;
            object tmpValues = worksheet.Range[0, 0, rowCount, columnCount].Value;
            end = DateTime.Now;

            getTime = (end - start).TotalSeconds; ;
            Console.WriteLine(string.Format("日付値（date）を取得する：{0:N3} 秒", getTime));

            start = DateTime.Now;
            workbook.Save("./files/gcexcel-saved-date.xlsx");
            end = DateTime.Now;
            saveTime = (end - start).TotalSeconds; ;
            Console.WriteLine(string.Format("日付値（date）を保存する：{0:N3} 秒", saveTime));

        }

        public static void TestSetRangeFormulas(int rowCount, int columnCount, ref double setTime, ref double calcTime, ref double saveTime, ref double usedMem)
        {

            Console.WriteLine();
            Console.WriteLine(string.Format("{0} 行 {1} 列で数式を使用する場合のベンチマーク", rowCount, columnCount));

            IWorkbook workbook = new Workbook();
            workbook.ReferenceStyle = ReferenceStyle.R1C1;
            IWorksheet worksheet = workbook.Worksheets[0];


            double[,] values = new double[rowCount, 2];

            for (int i = 0; i < rowCount; i++)
            {
                for (int j = 0; j < 2; j++)
                {
                    values[i, j] = i + j;
                }
            }
            worksheet.Range[0, 0, rowCount, 2].Value = values;

            DateTime start = DateTime.Now;
            worksheet.Range[0, 2, rowCount - 2, columnCount].Formula = "=SUM(RC[-2],RC[-1])";
            DateTime end = DateTime.Now;

            setTime = (end - start).TotalSeconds; ;
            Console.WriteLine(string.Format("数式を設定する：{0:N3} 秒", setTime));

            start = DateTime.Now;
            workbook.Calculate();
            end = DateTime.Now;

            calcTime = (end - start).TotalSeconds; ;
            Console.WriteLine(string.Format("数式を計算する：{0:N3} 秒", calcTime));

            workbook.ReferenceStyle = ReferenceStyle.A1;

            start = DateTime.Now;
            workbook.Save("./files/gcexcel-saved-formulas.xlsx");
            end = DateTime.Now;
            saveTime = (end - start).TotalSeconds; ;
            Console.WriteLine(string.Format("数式を保存する：{0:N3} 秒", saveTime));

        }


        public static void TestBigExcelFile(int rowCount, int columnCount, ref double openTime, ref double calcTime, ref double saveTime, ref double usedMem)
        {

            Console.WriteLine();
            Console.WriteLine(string.Format("多くの数値、数式、スタイルを含む大きなサイズのExcelファイルを使用する場合のベンチマーク"));

            IWorkbook workbook = new Workbook();

            DateTime start = DateTime.Now;
            workbook.Open("./files/test-performance.xlsx");
            DateTime end = DateTime.Now;

            openTime = (end - start).TotalSeconds; ;
            Console.WriteLine(string.Format("Excelファイルを開く：{0:N3} 秒", openTime));

            start = DateTime.Now;
            workbook.Dirty();
            workbook.Calculate();
            end = DateTime.Now;

            calcTime = (end - start).TotalSeconds; ;
            Console.WriteLine(string.Format("数式を計算する：{0:N3} 秒", calcTime));

            start = DateTime.Now;
            workbook.Save("./files/gcexcel-saved-test-performance.xlsx");
            end = DateTime.Now;
            saveTime = (end - start).TotalSeconds; ;
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

    internal static class DateTimeHelper
    {
        private static readonly System.DateTime Jan1st1970 = new System.DateTime(1970, 1, 1, 0, 0, 0, System.DateTimeKind.Utc);
        public static long CurrentUnixTimeMillis()
        {
            return (long)(System.DateTime.UtcNow - Jan1st1970).TotalSeconds;
        }
    }

    internal static class RectangularArrays
    {
        public static double[][] RectangularDoubleArray(int size1, int size2)
        {
            double[][] newArray = new double[size1][];
            for (int array1 = 0; array1 < size1; array1++)
            {
                newArray[array1] = new double[size2];
            }

            return newArray;
        }

        public static string[][] RectangularStringArray(int size1, int size2)
        {
            string[][] newArray = new string[size1][];
            for (int array1 = 0; array1 < size1; array1++)
            {
                newArray[array1] = new string[size2];
            }

            return newArray;
        }

        public static DateTime[][] RectangularDateTimeArray(int size1, int size2)
        {
            DateTime[][] newArray = new DateTime[size1][];
            for (int array1 = 0; array1 < size1; array1++)
            {
                newArray[array1] = new DateTime[size2];
            }

            return newArray;
        }
    }
}
