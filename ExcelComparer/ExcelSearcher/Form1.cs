using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelSearcher
{
    public partial class Form1 : Form
    {
        private delegate void DGSetProgress(float percent);
        private DGSetProgress SetProgress;

        private void SetProgressSafe(float percent)
        {
            if (this.InvokeRequired)
                this.Invoke(this.SetProgress, percent);
            else
            {
                toolStripProgressBar1.Value = (int) (percent * 100);
                toolStripProgressBar1.Visible = percent < 1;
            }
        }

        private delegate void DGUpdateTextboxSource(AutoCompleteStringCollection src);
        private DGUpdateTextboxSource UpdateTextboxSource;

        private void UpdateTextboxSourceSafe(AutoCompleteStringCollection src)
        {
            if (this.InvokeRequired)
                this.Invoke(this.UpdateTextboxSource, src);
            else
                textBox1.AutoCompleteCustomSource = src;
        }

        private delegate void DGUpdateStatueBarText(string txt);
        private DGUpdateStatueBarText UpdateStatueBarText;

        private void UpdateStatueBarTextSafe(string txt)
        {
            if (this.InvokeRequired)
                this.Invoke(this.UpdateStatueBarText, txt);
            else
            {
                toolStripStatusLabel1.Text = string.Format("{0} {1}", DateTime.Now, txt);
                statusStrip1.Invalidate();
            }
        }

        private string excelPath = "";
        private Thread rebuildThread;

        public Form1()
        {
            InitializeComponent();

            SetProgress = new DGSetProgress(a =>
            {
                toolStripProgressBar1.Value = (int) (a * 100);
                toolStripProgressBar1.Visible = a < 1;
            }); 
            UpdateTextboxSource = new DGUpdateTextboxSource(a => textBox1.AutoCompleteCustomSource = a);
            UpdateStatueBarText = new DGUpdateStatueBarText(a => toolStripStatusLabel1.Text = string.Format("{0} {1}", DateTime.Now, a));
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            FileAccessTimeManager.Load();
            if (File.Exists("config"))
            {
                using (var sr = new StreamReader("config"))
                {
                    excelPath = sr.ReadLine();
                }
            }
            KeywordManager.Init();
            textBox1.AutoCompleteCustomSource = KeywordManager.Get();

            rebuildThread = new Thread(CheckRebuild);
            rebuildThread.IsBackground = true;
            rebuildThread.Start();
            //richTextBoxEx1.SelectedText = "	2000	";
            //richTextBoxEx1.InsertLink("荆轲");
            //richTextBoxEx1.SelectedText = @"1	1	0	ProfilePictureTemporary	Character	ZhuJueHead	Character	{1,}	不详		Train/0000_train_data	2001			2001_YingZheng	2001_a	2001_a_prefab_weapon	黯然销魂刀	{100002,}	100002	20	{[0]={10,},[1]={20, }, [2]={30, }, [3]={40, }, [4]={50, }, [5]={60, }, }	{[0]=0,[1]=0,[2]=0,[3]=0,[4]=0,[5]=0,}	{}	{}	{}	{}	{}		{[0] = {[100002] = 10,},[1] = {[100002] = 10,},[2] = {[100002] = 10,},[3] = {[100002] = 10,},[4] = {[100002] = 10,},[5] = {[100002] = 10,},[6] = {[100002] = 10,},}	100002	{1,1}	100003";
        }

        private void RebuildCache()
        {
            if (!Directory.Exists("./cache"))
            {
                Directory.CreateDirectory("./cache");
            }

            if (string.IsNullOrEmpty(excelPath))
            {
                return;
            }
            var files = Directory.GetFiles(excelPath);


            List<string> toConvertList = new List<string>();
            var fileModifyDict = FileAccessTimeManager.FileModifyDict;
            foreach (var file in files)
            {
                var fileInfo = new FileInfo(file);
                if (file.EndsWith(".xlsx") && !fileInfo.Name.StartsWith("~"))
                {
                    if (!fileModifyDict.ContainsKey(fileInfo.Name) || fileModifyDict[fileInfo.Name] < TimeTool.DateTimeToUnixTime(fileInfo.LastWriteTime))
                        toConvertList.Add(file);
                }
            }

            if (toConvertList.Count > 0)
            {
                UpdateStatueBarTextSafe("开始重建缓存");
                var total = toConvertList.Count;
                int done = 0;
                SetProgressSafe(0);
                foreach (var file in toConvertList)
                {
                    var fileInfo = new FileInfo(file);
                    Excel2Csv.Convert(file, string.Format("./cache/{0}.txt", fileInfo.Name.Substring(0, fileInfo.Name.Length - 5)));
                    fileModifyDict[fileInfo.Name] = TimeTool.DateTimeToUnixTime(fileInfo.LastWriteTime);
                    done++;
                    SetProgressSafe((float)done / total);
                }
                FileAccessTimeManager.Save();
                SetProgressSafe(1);
                UpdateTextboxSourceSafe(KeywordManager.Get());
                UpdateStatueBarTextSafe("重建缓存完成");
            }
        }

        private void CheckRebuild()
        {
            while (true)
            {
                try
                {
                    RebuildCache();
                }
                catch (Exception)
                {
                }
                
                Thread.Sleep(5 * 1000);
            }
        }

        private void richTextBoxEx1_LinkClicked(object sender, System.Windows.Forms.LinkClickedEventArgs e)
        {
            var linkInfos = e.LinkText.Split('#');
            var fileInfos = linkInfos[1].Replace(".txt", "").Split('-');
            var path = string.Format("{0}/{1}.xlsx", excelPath, fileInfos[0]);
            string sheetName = fileInfos[1];//你的sheet的名字
            var cellName = ToExcelCellName(int.Parse(linkInfos[3]), int.Parse(linkInfos[2]));
            string strStart = cellName;//起始单元格
            string strEnd = cellName;//结束单元格
            object missing = Type.Missing;
            Excel.Application excel = new Excel.Application();
            Excel.Workbook book = excel.Workbooks.Open(path, missing,
                missing, missing, missing, missing, missing, missing, missing,
                missing, missing, missing, missing, missing, missing);
            Excel.Worksheet sheet = book.Worksheets[sheetName];
            excel.Application.Goto(sheet.Range[strStart, strEnd], true);
            excel.Visible = true;

            //System.Diagnostics.Process.Start("notepad.exe", "D:\debug\config.ini");
            //MessageBox.Show("A link has been clicked.\nThe link text is '" + e.LinkText + "'");
        }

        private void DoSearch()
        {
            UpdateStatueBarTextSafe("开始搜索");
            Invalidate();
            richTextBoxEx1.Clear();
            if (string.IsNullOrEmpty(excelPath))
            {
                richTextBoxEx1.SelectedText = "请先点击'绑定目录'选择绑定table文件夹";
                return;
            }
            int itemCount = 0;
            foreach (var file in Directory.GetFiles("./cache/"))
            {
                var fileInfo = new FileInfo(file);
                if (file.EndsWith(".txt"))
                {
                    using (StreamReader sr = new StreamReader(file, Encoding.UTF8))
                    {
                        string line;
                        int lineIndx = 0;
                        while ((line = sr.ReadLine()) != null)
                        {
                            var targetPos = line.IndexOf(textBox1.Text);
                            var sizeStr = textBox1.Text.Length;
                            bool isTarget = false;
                            int tabIndex = 0;
                            while (targetPos >= 0)
                            {
                                if (!isTarget)
                                {
                                    // 第一次进入先拼一个表头
                                    richTextBoxEx1.SelectionColor = Color.Red;
                                    var fileNames = fileInfo.Name.Replace(".txt", "").Split('-');
                                    if (fileNames[0].Length > 8)
                                        fileNames[0] = fileNames[0].Substring(0, 8);
                                    if (fileNames[1].Length > 8)
                                        fileNames[1] = fileNames[1].Substring(0, 8);
                                    richTextBoxEx1.SelectedText = string.Format("{0,-8}\t{1,-8}\t#{2:####}\t", fileNames[0], fileNames[1], lineIndx+14);
                                    richTextBoxEx1.SelectionColor = Color.Black;
                                }

                                if (targetPos > 0)
                                {
                                    var pickText = line.Substring(0, targetPos);
                                    richTextBoxEx1.SelectedText = pickText;
                                    tabIndex += GetAppearTimes(pickText, "\t");
                                }

                                richTextBoxEx1.InsertLink(line.Substring(targetPos, sizeStr), string.Format("{0}#{1}#{2}", fileInfo.Name, lineIndx + 14, tabIndex));

                                line = line.Substring(targetPos + sizeStr);
                                targetPos = line.IndexOf(textBox1.Text);
                                isTarget = true;
                            }

                            if (isTarget)
                            {
                                richTextBoxEx1.SelectedText = line + "\n";
                                itemCount++;
                                if (itemCount >= 1000)
                                {
                                    UpdateStatueBarTextSafe(string.Format("搜索结束:记录数达到上限，共找到记录{0}条", itemCount));
                                    //最多出1000条，抱歉
                                    return;
                                }
                            }

                            lineIndx++;
                        }
                    }
                }
            }

            UpdateStatueBarTextSafe(string.Format("搜索结束:共找到记录{0}条", itemCount));
            //防止link点击后乱跳转的bug
            this.richTextBoxEx1.Select(0, 0);
        }

        private void textBox1_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                DoSearch();
        }

        private void buttonSearch_Click(object sender, EventArgs e)
        {
            DoSearch();
        }

        int GetAppearTimes(string str1, string str2)
        {
            int i = 0;
            while (str1.IndexOf(str2) >= 0)
            {
                str1 = str1.Substring(str1.IndexOf(str2) + str2.Length);
                i++;
            }
            return i;
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

            return string.Join(string.Empty, chars.ToArray())+y;
        }

        private void buttonBind_Click(object sender, EventArgs e)
        {
            var openFileDialog = new FolderBrowserDialog();
            openFileDialog.Description = "请选择配表xlsx文件所在的目录";
            openFileDialog.ShowDialog();
            openFileDialog.ShowNewFolderButton = false;
            excelPath = openFileDialog.SelectedPath;
            if (!string.IsNullOrEmpty(excelPath))
            {
                using (var sw = new StreamWriter("config"))
                {
                    sw.Write(excelPath);
                }

                if (Directory.Exists("./cache/"))
                {//切目录时，cache清理掉
                    Directory.Delete("./cache/", true);
                }
                FileAccessTimeManager.FileModifyDict.Clear();
            }

            UpdateStatueBarTextSafe("绑定成功，请耐心等待重建缓存");
        }
    }
}
