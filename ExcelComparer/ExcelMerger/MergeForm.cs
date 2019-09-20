using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    public partial class MergeForm : Form
    {
        private bool hasErrorOnLoad;

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
                hasErrorOnLoad = true;
                MessageBox.Show(result, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }

            //MergeEvtData.Multiply(100);
            for (int i = 0; i < BaseMergeData.DataList.Count; i++)
            {
                var mergeData = BaseMergeData.DataList[i];
                mergeData.AddToDV(i, dataGridView1.Rows);
            }
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[1].Value == null)
                {
                    continue;
                }
                int index = int.Parse(row.Cells[1].Value.ToString());
                var mergeData = BaseMergeData.DataList[index];
                if (!mergeData.Conflict)
                {
                    HideButtons(row);
                }
            }
        }

        private void HideButtons(DataGridViewRow row)
        {
            row.Cells[6] = new DataGridViewTextBoxCell();
            row.Cells[6].Value = "";
            row.Cells[8] = new DataGridViewTextBoxCell();
            row.Cells[8].Value = "";

        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if(dataGridView1.Rows[e.RowIndex].Cells[1].Value == null)
                return;
            int index = int.Parse(dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString());
            var mergeData = BaseMergeData.DataList[index];
            if (e.ColumnIndex == 5 || e.ColumnIndex == 7)
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
            if (e.ColumnIndex == 6)
                useTheir = true;
            if (e.ColumnIndex == 8)
                useMine = true;

            var dt = BaseMergeData.DataList[e.RowIndex];
            dt.Resolve(useMine);
            HideButtons(dataGridView1.Rows[e.RowIndex]);
            dt.Conflict = false;
            dataGridView1.Invalidate();

            foreach (var evtData in BaseMergeData.DataList)
            {
                if (evtData.Conflict)
                    return;
            }

            Excel2Csv.OnLastErrorSolved();//行merge都在这里统一搞
            // Excel2Csv.Save();
        }

        private void MergeForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (hasErrorOnLoad)
            {
                //有错误，就让退吧
                return;
            }

            foreach (var evtData in BaseMergeData.DataList)
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
