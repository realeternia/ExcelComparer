using System;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelMerger
{
    public partial class MergeForm : Form
    {
        private bool hasErrorOnLoad;
        private bool hasSaved;

        public MergeForm()
        {
            InitializeComponent();
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToResizeRows = false;
        }

        private void MergeForm_Shown(object sender, EventArgs e)
        {
            var result = Excel2Csv.BeginMerge();
            if (result != "")
            {
                hasErrorOnLoad = true;
                MessageBox.Show(result, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
                return;
            }

            //MergeEvtData.Multiply(100);
            RefreshDvData();
        }

        private void RefreshDvData()
        {
            dataGridView1.Rows.Clear();
            for (int i = 0; i < BaseMergeData.DataList.Count; i++)
            {
                var mergeData = BaseMergeData.DataList[i];
                if (!mergeData.Conflict && toolStripButton1.Checked)
                {
                    continue;
                }

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
            if (dataGridView1.Rows[e.RowIndex].Cells[1].Value == null)
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
                        e.CellStyle.BackColor = mergeData.AutoMerge ? Color.Lime : Color.Orange;
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

            int index = int.Parse(dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString());
            var dt = BaseMergeData.DataList[index];
            dt.Resolve(useMine);
            // 选择因为可以来回切，所以不隐藏
            //  HideButtons(dataGridView1.Rows[e.RowIndex]);
            dt.Conflict = false;
            dataGridView1.Invalidate();
        }

        private void MergeForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (hasErrorOnLoad)
            {
                //有错误，就让退吧
                return;
            }

            if (!hasSaved)
            {
                if (MessageBox.Show("冲突解决未保存，点击“确定”继续关闭，结果不会保存。点击“取消”继续编辑", "警告", MessageBoxButtons.OKCancel,
                        MessageBoxIcon.Warning) == DialogResult.Cancel)
                {
                    e.Cancel = true;
                }

                return;
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            toolStripButton1.Checked = !toolStripButton1.Checked;
            RefreshDvData();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            foreach (var evtData in BaseMergeData.DataList)
            {
                if (evtData.Conflict)
                {
                    MessageBox.Show("还有未处理的冲突，无法保存", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }

            Excel2Csv.OnLastErrorSolved(); //行merge都在这里统一搞
            Excel2Csv.Save();
            hasSaved = true;
            toolStripButton2.Enabled = false;

            RunExeByProcess("svn.exe", "resolved " + ProArgs.Mine);

            MessageBox.Show("保存成功", "信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void RunExeByProcess(string exePath, string argument)
        {
            //开启新线程
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            //调用的exe的名称
            process.StartInfo.FileName = exePath;
            //传递进exe的参数
            process.StartInfo.Arguments = argument;
            process.StartInfo.UseShellExecute = false;
            //不显示exe的界面
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardInput = true;
            process.Start();

            process.StandardInput.AutoFlush = true;
            process.WaitForExit();
        }
    }
}
