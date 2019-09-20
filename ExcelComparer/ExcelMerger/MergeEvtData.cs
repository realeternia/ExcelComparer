using System.Collections.Generic;
using System.Windows.Forms;

namespace ExcelMerger
{
    public class MergeEvtData : BaseMergeData
    {
        public int Row;
        public int Column;
        public string OldValue;
        public string MyValue;
        public bool MyFormula;
        public string TheirsValue;
        public bool TheirFormula;

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

        public override void AddToDV(int i, DataGridViewRowCollection c)
        {
            c.Add(new string[] { i.ToString(), Label, OldValue, TheirsValue, "使用他的", MyValue, "使用我的" });
        }

        public override void Resolve(bool useMine)
        {
            if (useMine)
            {
                //本来就是用自己文件不用改
                // Excel2Csv.UpdAte(sheetName, dt.Row, dt.Column, dt.MyValue, dt.MyFormula);
                ConflictResult = MyValue;
            }
            else
            {
                var datas = Label.Split('-');
                var sheetName = datas[0];
                Excel2Csv.UpdAte(sheetName, Row, Column, TheirsValue, TheirFormula);
                ConflictResult = TheirsValue;
            }
        }
    }
}