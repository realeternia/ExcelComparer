using System.Collections.Generic;

namespace ExcelMerger
{
    public class MergeEvtData
    {
        public string Label;
        public int Row;
        public int Column;
        public string OldValue;
        public string MyValue;
        public bool MyFormula;
        public string TheirsValue;
        public bool TheirFormula;
        public bool Conflict;
        public string ConflictResult;

        public static List<MergeEvtData> DataList = new List<MergeEvtData>();

        public static void Add(string lb, int row, int column, string old, string my, bool myFormula, string their, bool theirFormula, bool conflict, string result)
        {
            MergeEvtData evtData = new MergeEvtData();
            evtData.Label = lb;
            evtData.Row = row;
            evtData.Column = column;
            evtData.OldValue = old;
            evtData.MyValue = my;
            evtData.MyFormula = myFormula;
            evtData.TheirsValue = their;
            evtData.TheirFormula = theirFormula;
            evtData.Conflict = conflict;
            evtData.ConflictResult = result;
            DataList.Add(evtData);
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
}