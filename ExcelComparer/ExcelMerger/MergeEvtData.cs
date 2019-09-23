using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using ExcelMerger.Properties;

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


        public static string ToExcelCellName(int x, int y)
        {
            if (x < 0) { throw new Exception("invalid parameter"); }

            List<string> chars = new List<string>();
            do
            {
                if (chars.Count > 0) x--;
                chars.Insert(0, ((char)(x % 26 + (int)'A')).ToString());
                x = (int)((x - x % 26) / 26);
            } while (x > 0);

            return y+string.Join(string.Empty, chars.ToArray());
        }

        public override void AddToDV(int i, DataGridViewRowCollection c)
        {
            c.Add(new object[] { Resources.warn, i.ToString(), Label, ToExcelCellName(Column-1, Row), OldValue, TheirsValue, "使用他的", MyValue, "保留我的" });
            c[c.Count - 1].Cells[0].ToolTipText = "格子数据发生修改";
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