using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace ExcelComparer
{
    class Program
    {
        private static string csvOldFile = "";
        private static string csvNewFile = "";

        public static float oldProgress;
        public static float newProgress;

        static void Main(string[] args)
        {
            var tortoiseMergePos = args[0];
            var oldPath = args[1];
            var newPath = args[2];

            ThreadStart method = () => TaskRun(oldPath, true);
            Thread t1 = new Thread(method);
            t1.IsBackground = true;
            t1.Start();

            ThreadStart method2 = () => TaskRun(newPath, false);
            Thread t2 = new Thread(method2);
            t2.IsBackground = true;
            t2.Start();

            Console.WriteLine("开始解析文件");
            while (csvOldFile == "" || csvNewFile == "")
            {
                Console.Write("\b\b\b\b\b\b\b\b\b\b\b\b\b\b\b\b对比进度{0:00.00}%", oldProgress * 50 + newProgress * 50);
                Thread.Sleep(10);
            }

            t1.Join();
            t2.Join();

            System.Diagnostics.Process.Start(tortoiseMergePos, string.Format("{0} {1}", csvOldFile, csvNewFile));
        }

        static void TaskRun(string path, bool isOld)
        {
            var name = Excel2Csv.Convert(path, isOld);
            if (isOld)
            {
                csvOldFile = name;
            }
            else
            {
                csvNewFile = name;
            }
        }
    }
}
