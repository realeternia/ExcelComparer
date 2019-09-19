using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelMerger
{
    public partial class MergeForm : Form
    {
        public MergeForm()
        {
            InitializeComponent();
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
        }

        private void MergeForm_Shown(object sender, EventArgs e)
        {
            Excel2Csv.BeginMerge();

            //MergeEvtData.Multiply(100);
            for (int i = 0; i < MergeEvtData.DataList.Count; i++)
            {
                var mergeData = MergeEvtData.DataList[i];
                dataGridView1.Rows.Add(new string[] { i.ToString(), mergeData.Label, mergeData.OldValue, mergeData.TheirsValue, "使用他的", mergeData.MyValue, "使用我的" });
            }
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[0].Value == null)
                {
                    continue;
                }
                int index = int.Parse(row.Cells[0].Value.ToString());
                var mergeData = MergeEvtData.DataList[index];
                if (!mergeData.Conflict)
                {
                    HideButtons(row);
                }
            }
        }

        private void HideButtons(DataGridViewRow row)
        {
            row.Cells[4] = new DataGridViewTextBoxCell();
            row.Cells[4].Value = "";
            row.Cells[6] = new DataGridViewTextBoxCell();
            row.Cells[6].Value = "";

        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if(dataGridView1.Rows[e.RowIndex].Cells[0].Value == null)
                return;
            int index = int.Parse(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
            var mergeData = MergeEvtData.DataList[index];
            if (e.ColumnIndex == 3 || e.ColumnIndex == 5)
            {
                if (mergeData.Conflict)
                {
                    e.CellStyle.BackColor = Color.LightCoral;
                }
                else
                {
                    var myVal = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
                    if (myVal == mergeData.ConflictResult)
                    {
                        e.CellStyle.BackColor = Color.Lime;
                    }
                }
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            bool useTheir = false;
            bool useMine = false;
            if (e.ColumnIndex == 4)
                useTheir = true;
            if (e.ColumnIndex == 6)
                useMine = true;

            var dt = MergeEvtData.DataList[e.RowIndex];
            var datas = dt.Label.Split('-');
            var sheetName = datas[0];
            if (useMine)
            {
                //本来就是用自己文件不用改
               // Excel2Csv.UpdAte(sheetName, dt.Row, dt.Column, dt.MyValue, dt.MyFormula);
               dt.ConflictResult = dt.MyValue;
            }
            else if (useTheir)
            {
                Excel2Csv.UpdAte(sheetName, dt.Row, dt.Column, dt.TheirsValue, dt.TheirFormula);
                dt.ConflictResult = dt.TheirsValue;
            }
            HideButtons(dataGridView1.Rows[e.RowIndex]);
            dt.Conflict = false;
            dataGridView1.Invalidate();
            // Excel2Csv.Save();
        }

        private void MergeForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Excel2Csv.CleanUp();
        }
    }
}
