using System.Collections.Generic;
using System.Deployment.Application;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelMerger
{
    public class Excel2Csv
    {
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

        public static void CleanUp()
        {
            if (epMine != null)
            {
                epMine.Save();
                epMine.Dispose();
                epMine = null;
            }
        }

        public static void BeginMerge()
        {
            var dtBase = LoadData(ProArgs.Base);
            var dtTheirs = LoadData(ProArgs.Theirs);

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
                                var baseCell = dtBase[idKey][col];
                                var theirsCell = dtTheirs[idKey][col];
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
        }


        private static Dictionary<string, List<CellData>> LoadData(string fileName)
        {
            var dict = new Dictionary<string, List<CellData>>();

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

                        dict[string.Format("{0}-key={1}", sheetIn.Name, dts[2].Content)] = dts;
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
            if (txt == "")
                return dc;

            return System.Drawing.Color.FromArgb(
                int.Parse(txt.Substring(2, 2), System.Globalization.NumberStyles.AllowHexSpecifier),
            int.Parse(txt.Substring(4, 2), System.Globalization.NumberStyles.AllowHexSpecifier),
            int.Parse(txt.Substring(6, 2), System.Globalization.NumberStyles.AllowHexSpecifier)
                );
        }
    }
}
