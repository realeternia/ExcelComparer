using System.Collections.Generic;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelMerger
{
    public class Excel2Csv
    {
        public class MyRowData
        {
            public string Tag;
            public List<CellData> Datas;

            public bool Eq(MyRowData other)
            {
                if (other == null || other.Datas == null || other.Datas.Count != Datas.Count)
                {
                    return false;
                }

                for (int i = 0; i < Datas.Count; i++)
                {
                    if (other.Datas[i].Content != Datas[i].Content)
                    {
                        return false;
                    }
                }

                return true;
            }
        }
        public struct CellData
        {
            public string Content;
            public bool IsFormula;

            public override string ToString()
            {
                return string.Format("{0} {1}", Content, IsFormula);
            }
        }


        private static ExcelPackage epMine;
        private static Dictionary<string, MyRowData> dtBase;
        private static Dictionary<string, MyRowData> dtTheirs;
        private static Dictionary<string, MyRowData> dtMine;

        public static void CleanUp()
        {
            if (epMine != null)
            {
                epMine.Save();
                epMine.Dispose();
                epMine = null;
            }
        }

        public static string BeginMerge()
        {
            dtBase = LoadData(ProArgs.Base, SheetMetaManager.AddBase);
            dtTheirs = LoadData(ProArgs.Theirs, SheetMetaManager.AddTheir);
            dtMine = LoadData(ProArgs.Mine, SheetMetaManager.AddMine); //这一次读文件其实可以省掉，但代码更难写
            MarkRowState(dtTheirs, dtBase, dtMine);
            MarkRowState(dtMine, dtBase, dtTheirs);

            foreach (var myRowData in dtMine)
            {
                if (myRowData.Value.Tag == "Add")
                {
                    if (dtTheirs[myRowData.Key].Tag == "Add")
                    {
                        bool sameData = true;
                        for (int j = 1; j < dtTheirs[myRowData.Key].Datas.Count; j++)
                        {
                            if (dtMine[myRowData.Key].Datas[j].Content != dtTheirs[myRowData.Key].Datas[j].Content)
                            {
                                sameData = false;
                                break;
                            }
                        }

                        if (sameData) //插了一样的数据，自动merge
                        {
                            MergeRowData.Add(myRowData.Key, "无记录", "添加了记录（相同）", "添加了记录（相同）", "", "", "添加了记录（相同）");
                        }
                        else
                        {
                            MergeRowData.Add(myRowData.Key, "无记录", "添加了记录", "添加了记录（我）", "Modify", "", "");
                        }
                    }
                }
                else if (myRowData.Value.Tag == "Delete")
                {
                    var otherTag = dtTheirs[myRowData.Key].Tag;
                    if (otherTag == "Modify")
                    {
                        MergeRowData.Add(myRowData.Key, "有记录", "修改了记录", "删除了记录（我）", "Add", "", "");
                    }
                }
                else if (myRowData.Value.Tag == "Modify")
                {
                    var otherTag = dtTheirs[myRowData.Key].Tag;
                    if (otherTag == "Modify")
                    {
                        // 不用考虑
                    }
                    else if (otherTag == "Delete")
                    {
                        MergeRowData.Add(myRowData.Key, "有记录", "删除了记录", "修改了记录（我）", "Delete", "", "");
                    }
                }
                else
                {
                    var otherTag = dtTheirs[myRowData.Key].Tag;
                    if (otherTag == "Add")
                    {
                        MergeRowData.Add(myRowData.Key, "有记录", "添加了记录", "无修改（我）", "Add", "", "添加了记录");
                    }
                    else if (otherTag == "Delete")
                    {
                        MergeRowData.Add(myRowData.Key, "有记录", "删除了记录", "无修改（我）", "Delete", "", "删除了记录");
                    }
                }
            }

            var compareResult = SheetMetaManager.CompareBaseAndTheir();
            if (compareResult != "")
            {
                return compareResult;
            }

            var fi = new FileInfo(ProArgs.Mine);

            //这里不using，因为需要hold住这个句柄，可能需要resolve冲突
            epMine = ExcelFileOpener.Open(fi, true);
            {
                var workbook = epMine.Workbook;

                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    var sheetIn = workbook.Worksheets[i];
                    int colCount = sheetIn.Dimension.End.Column;
                    for (int col = 2; col <= colCount; col++)
                    {
                        var dt = sheetIn.GetValue(9, col);
                        if (dt == null)
                        {
                            colCount = col;
                            break;
                        }
                    }

                    var compareResult2 = SheetMetaManager.AddAndCompareMine(i-1, sheetIn.Name, colCount);
                    if (compareResult2 != "")
                    {
                        return compareResult2;
                    }

                    for (int row = 14; row <= sheetIn.Dimension.End.Row; row++)
                    {
                        if (row >= 14 && sheetIn.GetValue(row, 2) == null)
                        {
                            break;
                        }

                        int id = 0;
                        for (int col = 2; col <= colCount; col++)
                        {
                            var cell = sheetIn.Cells[row, col];
                            if (cell != null)
                            {
                                if (col == 2)
                                {
                                    id = int.Parse(cell.Text);
                                }
                                var myCellContent = cell.Text;
                                bool myFormula = false;
                                if (!string.IsNullOrEmpty(cell.Formula))
                                {
                                    myCellContent = cell.Formula;
                                    myFormula = true;
                                }

                                string idKey = string.Format("{0}-key={1}", sheetIn.Name, id);
                                if (!dtTheirs.ContainsKey(idKey) || dtTheirs[idKey].Tag == "Add" || dtTheirs[idKey].Tag == "Delete" || dtMine[idKey].Tag == "Add")
                                {
                                    // 增减行的情况，这里不处理
                                    break;
                                }
                                var baseCell = dtBase[idKey].Datas[col];
                                var theirsCell = dtTheirs[idKey].Datas[col];

                                if (myCellContent != baseCell.Content || myCellContent != theirsCell.Content)
                                {
                                    bool conflict = false;
                                    string conflictResult = "";
                                    if (myCellContent == baseCell.Content && baseCell.Content != theirsCell.Content)
                                    {
                                        // 自动解决冲突，用别人的值
                                        UpdateInner(sheetIn, row, col, theirsCell.Content, theirsCell.IsFormula);
                                        conflictResult = theirsCell.Content;
                                    }
                                    else if (myCellContent != baseCell.Content && baseCell.Content != theirsCell.Content && myCellContent != theirsCell.Content)
                                    {
                                        conflict = true;
                                    }
                                    else
                                    {
                                        // 自动解决冲突，用自己的值
                                        conflictResult = myCellContent;
                                    }
                                    MergeEvtData.Add(idKey, row, col, baseCell.Content, myCellContent, myFormula, theirsCell.Content, theirsCell.IsFormula, conflict, conflictResult);
                                }
                            }
                        }
                    }
                }

            }
            return "";
        }

        private static void MarkRowState(Dictionary<string, MyRowData> dtTarget, Dictionary<string, MyRowData> dtBase, Dictionary<string, MyRowData> dtOther)
        {
            foreach (var theirRow in dtTarget)
            {
                if (!dtBase.ContainsKey(theirRow.Key))
                {
                    if (theirRow.Value.Datas != null)
                        theirRow.Value.Tag = "Add";
                    if (!dtOther.ContainsKey(theirRow.Key))
                    {
                        dtOther[theirRow.Key] = new MyRowData { Tag = "" }; //帮另一组加一条
                    }
                }
                else if (!dtBase[theirRow.Key].Eq(theirRow.Value))
                {
                    theirRow.Value.Tag = "Modify";
                }
            }

            foreach (var baseRow in dtBase)
            {
                if (!dtTarget.ContainsKey(baseRow.Key) || dtTarget[baseRow.Key].Datas == null)
                {
                    dtTarget[baseRow.Key] = new MyRowData {Tag = "Delete"};
                }
            }
        }


        private static Dictionary<string, MyRowData> LoadData(string fileName, SheetMetaManager.AddData func)
        {
            var dict = new Dictionary<string, MyRowData>();

            var fi = new FileInfo(fileName);
            using (ExcelPackage ep = ExcelFileOpener.Open(fi, false))
            {
                var workbook = ep.Workbook;

                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    var sheetIn = workbook.Worksheets[i];
                    int colCount = sheetIn.Dimension.End.Column;
                    for (int col = 2; col <= colCount; col++)
                    {
                        var dt = sheetIn.GetValue(9, col);
                        if (dt == null)
                        {
                            colCount = col;
                            break;
                        }
                    }

                    bool isRegular = colCount > 0 && sheetIn.Dimension.End.Row >= 13 &&
                                     sheetIn.Cells[13, 1].Text == "BEGIN";
                    if (!isRegular)
                    {
                        continue;
                    }
                    func(sheetIn.Name, colCount);

                    for (int row = 14; row <= sheetIn.Dimension.End.Row; row++)
                    {
                        if (row >= 14 && sheetIn.GetValue(row, 2) == null)
                        {
                            break;
                        }

                        var dts = new List<CellData>();
                        dts.Add(new CellData()); //excel下标从1开始
                        for (int col = 1; col <= colCount; col++)
                        {
                            CellData cellInfo = new CellData();
                            var cell = sheetIn.Cells[row, col];
                            if (cell != null)
                            {
                                if (!string.IsNullOrEmpty(cell.Formula))
                                {
                                    cellInfo.Content = cell.Formula;
                                    cellInfo.IsFormula = true;
                                }
                                else
                                {
                                    cellInfo.Content = cell.Text;
                                }
                            }
                            dts.Add(cellInfo);
                        }

                        var rowData = new MyRowData();
                        rowData.Datas = dts;
                        dict[string.Format("{0}-key={1}", sheetIn.Name, dts[2].Content)] = rowData;
                    }
                }
            }

            return dict;
        }

        private static void UpdateInner(ExcelWorksheet sheetIn, int row, int col, string val, bool isFormula)
        {
            var targetCell = sheetIn.Cells[row, col];
            var stl = targetCell.Style;
            targetCell.Clear(); //单元格内部多种格式需要清掉，不然会出现赋值不了的情况
            if (stl.Font != null)
            {
                targetCell.Style.Font.Name = stl.Font.Name;
                targetCell.Style.Font.Size = stl.Font.Size;
                targetCell.Style.Font.Italic = stl.Font.Italic;
                targetCell.Style.Font.Bold = stl.Font.Bold;
                targetCell.Style.Font.Color.SetColor(ParseColor(stl.Font.Color.Rgb, Color.Black));
            }
            if (stl.Fill != null)
            {
                targetCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                targetCell.Style.Fill.BackgroundColor.SetColor(ParseColor(stl.Fill.BackgroundColor.Rgb, Color.White));
            }

            // 直接赋值好像不对，只能像上面一个一个属性赋值了
            //targetCell.Style.Fill = stl.Fill;
            if (isFormula)
            {
                targetCell.Formula = val;
                sheetIn.SetValue(row, col, targetCell);
            }
            else
            {
                targetCell.Value = val;
            }
        }

        public static void UpdAte(string sheetName, int row, int col, string val, bool isFormula)
        {
            ExcelWorksheet sheet = null;
            foreach (var workbookWorksheet in epMine.Workbook.Worksheets)
            {
                if (workbookWorksheet.Name == sheetName)
                {
                    sheet = workbookWorksheet;
                    break;
                }
            }

            if (sheet == null)
            {
                return;
            }

            UpdateInner(sheet, row, col, val, isFormula);
        }

        private static Color ParseColor(string txt, Color dc)
        {
            if (string.IsNullOrEmpty(txt))
                return dc;

            return System.Drawing.Color.FromArgb(
                int.Parse(txt.Substring(2, 2), System.Globalization.NumberStyles.AllowHexSpecifier),
            int.Parse(txt.Substring(4, 2), System.Globalization.NumberStyles.AllowHexSpecifier),
            int.Parse(txt.Substring(6, 2), System.Globalization.NumberStyles.AllowHexSpecifier)
                );
        }

        public static void OnLastErrorSolved()
        {
            Dictionary<string ,MyRowData> addList = new Dictionary<string, MyRowData>();
            List<string> removeList = new List<string>();
            Dictionary<string, MyRowData> modifyList = new Dictionary<string, MyRowData>();
            foreach (var baseMergeData in BaseMergeData.DataList)
            {
                if (!baseMergeData.IsRowError)
                    continue;

                var rowError = baseMergeData as MergeRowData;

                //使用我的修改就忽略掉
                if (rowError.ConflictResult == rowError.DesMine)
                    continue;

                if (rowError.TheirTag == "Add")
                {
                    addList[rowError.Label] = dtTheirs[rowError.Label];
                }
                else if (rowError.TheirTag == "Delete")
                {
                    removeList.Add(rowError.Label);
                }
                else if (rowError.TheirTag == "Modify")
                {
                    modifyList[rowError.Label] = dtTheirs[rowError.Label];
                }
            }

            var workbook = epMine.Workbook;

            for (int i = 1; i <= workbook.Worksheets.Count; i++)
            {
                var sheetIn = workbook.Worksheets[i];
                int row;
                for (row = 14; row <= sheetIn.Dimension.End.Row; row++)
                {
                    var idCellStr = sheetIn.GetValue(row, 2);
                    if (row >= 14 && idCellStr == null)
                        break;

                    string rowKey = string.Format("{0}-key={1}", sheetIn.Name, idCellStr.ToString());
                    if (removeList.Contains(rowKey))
                    {
                        sheetIn.DeleteRow(row);
                        row--;
                        continue;
                    }

                    MyRowData rowData;
                    if (modifyList.TryGetValue(rowKey, out rowData))
                    {
                        //做替换
                        for (int j = 1; j < rowData.Datas.Count; j++)
                        {
                            UpdateInner(sheetIn, row, j, rowData.Datas[j].Content, rowData.Datas[j].IsFormula);
                        }
                    }
                }

                foreach (var rowData in addList)
                {
                    if (rowData.Key.StartsWith(sheetIn.Name))
                    {
                        sheetIn.InsertRow(row, 1, row-1);
                        row++; //插入到最后
                        for (int j = 1; j < rowData.Value.Datas.Count; j++)
                        {
                            UpdateInner(sheetIn, row, j, rowData.Value.Datas[j].Content, rowData.Value.Datas[j].IsFormula);
                        }
                    }
                }
            }

        }
    }
}
