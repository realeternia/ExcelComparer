using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace ExcelSearcher
{
    public class FileAccessTimeManager
    {
        public static Dictionary<string, uint> FileModifyDict = new Dictionary<string, uint>();

        public static void Save()
        {
            using (var sw = new StreamWriter("./cache/time.dt"))
            {
                foreach (var u in FileModifyDict)
                {
                    sw.WriteLine("{0}\t{1}", u.Key, u.Value);
                }
            }
        }

        public static void Load()
        {
            if (!File.Exists("./cache/time.dt"))
            {
                return;
            }
            using (var sr = new StreamReader("./cache/time.dt"))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    var dts = line.Split('\t');
                    if (dts.Length >= 2)
                        FileModifyDict[dts[0]] = uint.Parse(dts[1]);
                }
            }
        }
    }
}