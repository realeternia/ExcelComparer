using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    public partial class MergeForm : Form
    {
        private bool hasError;

        public MergeForm()
        {
            InitializeComponent();
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
        }

        private void MergeForm_Shown(object sender, EventArgs e)
        {
            var result = Excel2Csv.BeginMerge();
            if (result != "")
            {
                hasError = true;
                MessageBox.Show(result, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }

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
            if (hasError)
            {
                //有错误，就让退吧
                return;
            }

            foreach (var evtData in MergeEvtData.DataList)
            {
                if (evtData.Conflict)
                {
                    if (MessageBox.Show("还有未处理的冲突，如果关闭解决过程不会保存。点击“取消”继续编辑", "警告", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.Cancel)
                    {
                        e.Cancel = true;
                    }
                    return;
                }
            }

            Excel2Csv.CleanUp();
        }
    }
}
