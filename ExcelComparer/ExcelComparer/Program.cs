using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelComparer
{
    class Program
    {
        static void Main(string[] args)
        {
            var tortoiseMergePos = args[0];
            var oldPath = args[1];
            var newPath = args[2];

            var csvOldFile = Excel2Csv.Convert(oldPath);
            var csvNewFile = Excel2Csv.Convert(newPath);

            System.Diagnostics.Process.Start(tortoiseMergePos, string.Format("{0} {1}", csvOldFile, csvNewFile));
        }
    }
}
