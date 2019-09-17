using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace ExcelComparer
{
    public class Excel2Csv
    {
        public static string Convert(string fileName, bool isOld)
        {
            var fi = new FileInfo(fileName);
            var tempFile = Path.GetTempFileName();
            StreamWriter sw = new StreamWriter(tempFile);
            using (ExcelPackage ep =  ExcelFileOpener.Open(fi, false))
            {
                var workbook = ep.Workbook;

                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    WriteSheet(workbook.Worksheets[i], sw, isOld, i, workbook.Worksheets.Count);
                }
            }

            sw.Close();
            return tempFile;
        }

        private static void WriteSheet(ExcelWorksheet sheetIn, StreamWriter sw, bool isOld, int now, int total)
        {
            if (sheetIn.Name.StartsWith("~") || sheetIn.Dimension == null)
                return;

            int colCount = sheetIn.Dimension.End.Column;
            for (int col = 2; col <= colCount; col++)
            {
                var dt = sheetIn.GetValue(9, col);
                if (dt == null)
                {
                    colCount = col;
                    break;
                }
            }

            for (int row = 1; row <= sheetIn.Dimension.End.Row; row++)
            {
                if (row > 14 && sheetIn.GetValue(row, 2) == null)
                {
                    break;
                }

                for (int col = 1; col <= colCount; col++)
                {
                    sw.Write(sheetIn.GetValue(row, col) + "\t"); 
                }

                sw.WriteLine();

                var progress = ((float) (now-1) / total) + (float)row / sheetIn.Dimension.End.Row / total;
                if (isOld)
                {
                    Program.oldProgress = progress;
                }
                else
                {
                    Program.newProgress = progress;
                }
            }
        }

    }
}
