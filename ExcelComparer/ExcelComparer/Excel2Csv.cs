using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace ExcelComparer
{
    public class Excel2Csv
    {
        public static string Convert(string fileName)
        {
            var fi = new FileInfo(fileName);
            var tempFile = Path.GetTempFileName();
            StreamWriter sw = new StreamWriter(tempFile);
            using (ExcelPackage ep =  ExcelFileOpener.Open(fi, false))
            {
                var workbook = ep.Workbook;

                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    WriteSheet(workbook.Worksheets[i], sw);
                }
            }

            sw.Close();
            return tempFile;
        }

        private static void WriteSheet(ExcelWorksheet sheetIn, StreamWriter sw)
        {
            if (sheetIn.Name.StartsWith("~") || sheetIn.Dimension == null)
                return;

            for (int row = 1; row <= sheetIn.Dimension.End.Row; row++)
            {
                for (int col = 1; col <= sheetIn.Dimension.End.Column; col++)
                {
                    sw.Write(sheetIn.GetValue(row, col) + "\t"); 
                }

                sw.WriteLine();
            }
        }

    }
}
