using System.Collections.Generic;
using System.Windows.Forms;

namespace ExcelMerger
{
    public class BaseMergeData
    {
        public static List<BaseMergeData> DataList = new List<BaseMergeData>();

        public string Label; // sheet + key
        public string ConflictResult;
        public bool Conflict;
        public virtual bool IsRowError
        {
            get { return false; }
        }

        public virtual void AddToDV(int i, DataGridViewRowCollection c)
        {
        }

        public virtual void Resolve(bool useMine)
        {

        }

        public static void Multiply(int x)
        {
            // 压力测试用
            var dts = DataList.ToArray();
            for (int i = 0; i < x; i++)
            {
                DataList.AddRange(dts);
            }
        }
    }

    public class MergeRowData : BaseMergeData
    {
        public string DesBase;
        public string DesTheir;
        public string DesMine;
        public string TheirTag;
        public string MyTag;
        public override bool IsRowError
        {
            get { return true; }
        }

        public static void Add(string lb, string dbase, string dtheir, string dmine, string ttheir, string tmine, string result)
        {
            MergeRowData evtData = new MergeRowData();
            evtData.Label = lb;
            evtData.DesBase = dbase;
            evtData.DesTheir = dtheir;
            evtData.DesMine = dmine;
            evtData.TheirTag = ttheir;
            evtData.MyTag = tmine;
            evtData.Conflict = result == "";
            evtData.ConflictResult = result;
            DataList.Add(evtData);
        }

        public override void AddToDV(int i, DataGridViewRowCollection c)
        {
            c.Add(new string[] { i.ToString(), Label, DesBase, DesTheir, "使用他的", DesMine, "使用我的" });
        }
        public override void Resolve(bool useMine)
        {
            if (useMine)
            {
                ConflictResult = DesMine;
            }
            else
            {
                ConflictResult = DesTheir;
            }
        }
    }
}