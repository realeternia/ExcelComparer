using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace ExcelSearcher
{
    public class KeywordManager
    {
        private static Dictionary<string, bool> keywordDict = new Dictionary<string, bool>();

        public static void Init()
        {
            if (!Directory.Exists("./cache/"))
            {
                return;
            }
            foreach (var file in Directory.GetFiles("./cache/"))
            {
                if (file.EndsWith(".txt"))
                {
                    using (StreamReader sr = new StreamReader(file, Encoding.UTF8))
                    {
                        string line;
                        while ((line = sr.ReadLine()) != null)
                        {
                            var dts = line.Split('\t');
                            int c;
                            foreach (var dt in dts)
                            {
                                if (!int.TryParse(dt, out c) && dt.Length > 1 && dt.Length <= 5)
                                {
                                    //数字不要，字符长度大于5的不要
                                    KeywordManager.SetKeyword(dt);
                                }
                            }
                        }
                    }
                }
            }
        }

        public static void SetKeyword(string wd)
        {
            keywordDict[wd] = true;
        }

        public static AutoCompleteStringCollection Get()
        {
            var source = new AutoCompleteStringCollection();
            foreach (var pair in keywordDict)
                source.Add(pair.Key);
            return source;
        }
    }
}