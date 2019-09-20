using System.Collections.Generic;

namespace ExcelMerger
{
    public class SheetMetaManager
    {
        public delegate void AddData(string name, int column);

        public struct SheetMetaData
        {
            public string SheetName;
            // 列数，列数互相之间不一样也不合并
            public int ColumnCount;
        }

        private static List<SheetMetaData> dictBase = new List<SheetMetaData>();
        private static List<SheetMetaData> dictTheir = new List<SheetMetaData>();
        private static List<SheetMetaData> dictMine = new List<SheetMetaData>();

        public static void AddBase(string name, int column)
        {
            dictBase.Add(new SheetMetaData {SheetName = name, ColumnCount = column});
        }
        public static void AddTheir(string name, int column)
        {
            dictTheir.Add(new SheetMetaData { SheetName = name, ColumnCount = column });
        }
        public static void AddMine(string name, int column)
        {
            dictMine.Add(new SheetMetaData { SheetName = name, ColumnCount = column });
        }

        public static string CompareBaseAndTheir()
        {
            if (dictBase.Count == 0)
            {
                return "该文件没有可以识别的分页，不需要合并";
            }

            if (dictBase.Count != dictTheir.Count)
            {
                return "别人修改了文件的sheet数量，无法继续合并";
            }

            for (int i = 0; i < dictBase.Count; i++)
            {
                if (dictBase[i].SheetName != dictTheir[i].SheetName)
                {
                    return "别人修改了文件的sheet标题或调整了顺序，无法继续合并";
                }
                if (dictBase[i].ColumnCount != dictTheir[i].ColumnCount)
                {
                    return string.Format("别人新建或删除了sheet[{0}]的列，无法继续合并", dictTheir[i].SheetName);
                }
            }

            return "";
        }

        public static string AddAndCompareMine(int idx, string name, int column)
        {
            if (idx >= dictBase.Count)
            {
                return "你修改了文件的sheet数量，无法继续合并";
            }

            if (dictBase[idx].SheetName != name)
            {
                return "你修改了文件的sheet标题或调整了顺序，无法继续合并";
            }
            if (dictBase[idx].ColumnCount != column)
            {
                return string.Format("你新建或删除了sheet[{0}]的列，无法继续合并", dictTheir[idx].SheetName);
            }

            return "";
        }
    }
}