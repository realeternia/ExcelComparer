using System.IO;
using OfficeOpenXml;

namespace ExcelSearcher
{
    public class Excel2Csv
    {
        public static void Convert(string fileName, string outName)
        {
            var fi = new FileInfo(fileName);
            
            using (ExcelPackage ep =  ExcelFileOpener.Open(fi, false))
            {
                var workbook = ep.Workbook;

                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    var name = string.Format("{0}-{1}.txt", outName, workbook.Worksheets[i].Name);
                    StreamWriter sw = new StreamWriter(name);
                    WriteSheet(workbook.Worksheets[i], sw);
                    sw.Close();
                }
            }
        }

        private static void WriteSheet(ExcelWorksheet sheetIn, StreamWriter sw)
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
            for (int row = 14; row <= sheetIn.Dimension.End.Row; row++)
            {
                if (sheetIn.GetValue(row, 2) == null)
                {
                    break;
                }
                for (int col = 1; col <= colCount; col++)
                {
                    var cell = sheetIn.GetValue(row, col);
                    var r = "\t";
                    if (cell != null)
                    {
                        r = cell.ToString().Replace("\n", "") + "\t";
                    }

                    sw.Write(r); 
                }

                sw.WriteLine();
            }
        }

    }
}
