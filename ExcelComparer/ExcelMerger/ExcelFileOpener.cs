using System.IO;
using OfficeOpenXml;

namespace ExcelMerger
{
    internal class ExcelFileOpener
    {

        public static ExcelPackage Open(FileInfo file, bool isWrite)
        {
            if (isWrite)
            {
                return OpenWrite(file);
            }
            return OpenRead(file);
        }

        private static ExcelPackage OpenWrite(FileInfo file)
        {
            ExcelPackage ep = null;

            ep = new ExcelPackage(file);
            return ep;
        }

        private static ExcelPackage OpenRead(FileInfo file)
        {
            ExcelPackage ep = null;
            string tempFile = null;
            try
            {
                ep = new ExcelPackage(file);
            }
            catch
            {
                tempFile = Path.GetTempFileName();
                ep = new ExcelPackage(file.CopyTo(tempFile, true));
            }

            return ep;
        }
    }
}
