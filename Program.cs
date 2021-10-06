using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace GcExcelPerformanceTest
{
    class Program
    {
        static void Main(string[] args)
        {
            double setTime = 0, getTime = 0, saveTime = 0, usedMem = 0;
            int row = 100000, col = 30;
            GC.Collect();

            long memDiff = GetMaxTotalMemoryAllocation(() =>
            {
                GcExcelBenchmark.TestSetRangeValues_Double(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            });
            Console.WriteLine(string.Format("メモリ使用量：{0:N3} MB ", memDiff / 1024 / 1024));

            GC.Collect();

            memDiff = GetMaxTotalMemoryAllocation(() =>
            {
                GcExcelBenchmark.TestSetRangeValues_String(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            });
            Console.WriteLine(string.Format("メモリ使用量：{0:N3} MB ", memDiff / 1024 / 1024));

            GC.Collect();

            memDiff = GetMaxTotalMemoryAllocation(() =>
            {
                GcExcelBenchmark.TestSetRangeValues_Date(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            });
            Console.WriteLine(string.Format("メモリ使用量：{0:N3} MB ", memDiff / 1024 / 1024));

            GC.Collect();

            memDiff = GetMaxTotalMemoryAllocation(() =>
            {
                GcExcelBenchmark.TestSetRangeFormulas(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            });
            Console.WriteLine(string.Format("メモリ使用量：{0:N3} MB ", memDiff / 1024 / 1024));

            GC.Collect();

            memDiff = GetMaxTotalMemoryAllocation(() =>
            {
                GcExcelBenchmark.TestBigExcelFile(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            });
            Console.WriteLine(string.Format("メモリ使用量：{0:N3} MB ", memDiff / 1024 / 1024));

            GC.Collect();

            memDiff = GetMaxTotalMemoryAllocation(() =>
            {
                NPOIBenchmark.TestSetRangeValues_Double(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            });
            Console.WriteLine(string.Format("メモリ使用量：{0:N3} MB ", memDiff / 1024 / 1024));

            GC.Collect();

            memDiff = GetMaxTotalMemoryAllocation(() =>
            {
                NPOIBenchmark.TestSetRangeValues_String(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            });
            Console.WriteLine(string.Format("メモリ使用量：{0:N3} MB ", memDiff / 1024 / 1024));

            GC.Collect();

            memDiff = GetMaxTotalMemoryAllocation(() =>
            {
                NPOIBenchmark.TestSetRangeValues_Date(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);

            });
            Console.WriteLine(string.Format("メモリ使用量：{0:N3} MB ", memDiff / 1024 / 1024));

            GC.Collect();

            memDiff = GetMaxTotalMemoryAllocation(() =>
            {
                NPOIBenchmark.TestSetRangeFormulas(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            });
            Console.WriteLine(string.Format("メモリ使用量：{0:N3} MB ", memDiff / 1024 / 1024));

            GC.Collect();

            memDiff = GetMaxTotalMemoryAllocation(() =>
            {
                NPOIBenchmark.TestBigExcelFile(row, col, ref setTime, ref getTime, ref saveTime, ref usedMem);
            });
            Console.WriteLine(string.Format("メモリ使用量：{0:N3} MB ", memDiff / 1024 / 1024));

            Console.ReadKey();
        }

        private static long GetMaxTotalMemoryAllocation(Action testAction)
        {
            var initMemory = GC.GetTotalMemory(true);
            var memoryUsages = new System.Collections.Generic.List<long>();
            var cancelSource = new CancellationTokenSource();
            var cancelToken = cancelSource.Token;
            var memUsageCounter = Task.Run(async () =>
            {
                try
                {
                    while (true)
                    {
                        memoryUsages.Add(GC.GetTotalMemory(false));
                        await Task.Delay(100, cancelToken);
                    }
                }
                catch (TaskCanceledException)
                {
                }
            });
            testAction();
            cancelSource.Cancel();
            memoryUsages.Add(GC.GetTotalMemory(false));
            var highMem = memoryUsages.Max();
            var memDiff = highMem - initMemory;
            return memDiff;
        }


    }
}
