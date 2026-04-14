using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Microsoft.Office.Interop.Excel;
using static System.Net.Mime.MediaTypeNames;
using _Excel = Microsoft.Office.Interop.Excel;

namespace CommonFunctions
{
    class ExcelHelper
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook MyWorkbook;
        Worksheet MyWorksheet;

        public ExcelHelper()
        {

        }

        public ExcelHelper(string path)
        {

        }

        public ExcelHelper(string path, int Sheet)
        {
            this.path = path;
            MyWorkbook = excel.Workbooks.Open(path);
            MyWorksheet = MyWorkbook.Worksheets[Sheet];
        }

        public void CreateNewFile()
        {
            this.MyWorkbook = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
        }

        public void CreateNewSheet()
        {
            Worksheet NewSheet = MyWorkbook.Worksheets.Add(After: MyWorksheet);
        }

        public string ReadCell(int row, int column)
        {
            if (MyWorksheet.Cells[row, column].Value2 != null)
            {
                return MyWorksheet.Cells[row, column].Value2;
            }
            else { return ""; }
        }

        public string[,] ReadRangeString(int startRow, int startColumn, int endRow, int endColumn)
        {
            Range range = (Range)MyWorksheet.Range[MyWorksheet.Cells[startRow, startColumn], MyWorksheet.Cells[endRow, endColumn]];
            object[,] holder = range.Value2;
            string[,] returnstring = new string[(endRow - startRow) + 1, (endColumn - startColumn) + 1];
            for (int p = 1; p <= ((endRow - startRow) + 1); p++)
            {
                for (int q = 1; q <= ((endColumn - startColumn) + 1); q++)
                {
                    returnstring[p - 1, q - 1] = holder[p, q].ToString();
                }
            }
            return returnstring;
        }

        public double[,] ReadRangeDouble(int startRow, int startColumn, int endRow, int endColumn)
        {
            Range range = (Range)MyWorksheet.Range[MyWorksheet.Cells[startRow, startColumn], MyWorksheet.Cells[endRow, endColumn]];
            object[,] holder = range.Value2;
            double[,] returndouble = new double[(endRow - startRow) + 1, (endColumn - startColumn) + 1];
            for (int p = 1; p <= ((endRow - startRow) + 1); p++)
            {
                for (int q = 1; q <= ((endColumn - startColumn) + 1); q++)
                {
                    returndouble[p - 1, q - 1] = Convert.ToDouble(holder[p, q]);
                }
            }
            return returndouble;
        }

        public void WriteRangeString(int startRow, int startColumn, int endRow, int endColumn, string[,] writestring)
        {
            Range range = (Range)MyWorksheet.Range[MyWorksheet.Cells[startRow, startColumn], MyWorksheet.Cells[endRow, endColumn]];
            range.Value2 = writestring;
        }

        public void WriteRangeDouble(int startRow, int startColumn, int endRow, int endColumn, double[,] writestring)
        {
            Range range = (Range)MyWorksheet.Range[MyWorksheet.Cells[startRow, startColumn], MyWorksheet.Cells[endRow, endColumn]];
            range.Value2 = writestring;
        }

        public void WriteRangeArray(int startStroka, int Stolbec, double[] writearray)
        {
            Range range = (Range)MyWorksheet.Range[MyWorksheet.Cells[startStroka, Stolbec], MyWorksheet.Cells[((writearray.Length - 1) + startStroka), Stolbec]];
            range.Value2 = writearray;
        }

        public void WriteToCellString(int row, int column, string data)
        {
            MyWorksheet.Cells[row, column].Value2 = data;
        }

        public void WriteToCellInt(int row, int column, int data)
        {
            MyWorksheet.Cells[row, column].Value2 = data;
        }

        public void WriteToCellDouble(int row, int column, double data)
        {
            MyWorksheet.Cells[row, column].Value2 = data;
        }

        public void MergeCells(string Cell1, string Cell2)
        {
            Range Range = (Range)MyWorksheet.get_Range(Cell1, Cell2);
            Range.Merge(Type.Missing);
        }

        public void WrapTextInCells(string Cell1, string Cell2)
        {
            Range Range = (Range)MyWorksheet.get_Range(Cell1, Cell2);
            Range.WrapText = true;
        }

        public void ChangeColumnWidth(int Column, double Width)
        {
            MyWorksheet.Columns[Column].ColumnWidth = Width;
        }

        public void AutofitColumns(string Cell1, string Cell2)
        {
            Range aRange = MyWorksheet.get_Range(Cell1, Cell2);
            aRange.Columns.AutoFit();
        }

        public void CenterAligmentColumns(string Cell1, string Cell2)
        {
            Range aRange = MyWorksheet.get_Range(Cell1, Cell2);
            aRange.Columns.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            aRange.Columns.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
        }

        public void AddBordersToRange(string Cell1, string Cell2)
        {
            Range Range = MyWorksheet.get_Range(Cell1, Cell2);
            Range.Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
        }

        public void AddBoldBordersToRange(string Cell1, string Cell2)
        {
            Range Range = MyWorksheet.get_Range(Cell1, Cell2);
            Range.Cells.Borders.LineStyle = XlLineStyle.xlContinuous;
            Range.Cells.Borders.Weight = 3d;
        }

        public void AddBoldBordersAroundToRange(string Cell1, string Cell2)
        {
            Range Range = MyWorksheet.get_Range(Cell1, Cell2);
            Range.Cells.Borders[XlBordersIndex.xlEdgeLeft].Weight = 3d;
            Range.Cells.Borders[XlBordersIndex.xlEdgeRight].Weight = 3d;
            Range.Cells.Borders[XlBordersIndex.xlEdgeTop].Weight = 3d;
            Range.Cells.Borders[XlBordersIndex.xlEdgeBottom].Weight = 3d;
        }

        public void AddFormulaToCell(string Cell1, Object Formula)
        {
            Range Range = MyWorksheet.get_Range(Cell1);
            Range.Formula = Formula;
        }

        public void ChangeCellNumberFormat(string Cell, string NumberFormat)
        {
            Range Range = MyWorksheet.get_Range(Cell);
            Range.NumberFormat = NumberFormat;
        }

        public ChartObject CreateScatterGraph(int StolbecX, int startStrokaX, int endStrokaX, int StolbecY, int startStrokaY, int endStrokaY, string GraphName)
        {
            ChartObjects cbs = (ChartObjects)MyWorksheet.ChartObjects();
            ChartObject cb = cbs.Add(400, 7365, 700, 350);
            cb.Chart.HasTitle = true;
            cb.Chart.ChartTitle.Text = GraphName;
            cb.Chart.ChartType = XlChartType.xlXYScatterSmooth;
            cb.Chart.HasTitle = true;
            Range cRangeX = (Range)MyWorksheet.Range[MyWorksheet.Cells[startStrokaX, StolbecX], MyWorksheet.Cells[endStrokaX, StolbecX]];
            Microsoft.Office.Interop.Excel.SeriesCollection oSeriesCollection = (Microsoft.Office.Interop.Excel.SeriesCollection)cb.Chart.SeriesCollection();

            for (int i = 0; i < 1; i++)
            {
                Microsoft.Office.Interop.Excel.Series oSeries = oSeriesCollection.NewSeries();
                oSeries.XValues = cRangeX;
                oSeries.Values = (Range)MyWorksheet.Range[MyWorksheet.Cells[startStrokaY, StolbecY], MyWorksheet.Cells[endStrokaY, StolbecY]];
            }
            return cb;
        }

        public Trendline CreateLineTrendline(Microsoft.Office.Interop.Excel.Series oSeries)
        {
            Microsoft.Office.Interop.Excel.Trendline myTrend = ((Microsoft.Office.Interop.Excel.Trendlines)oSeries.Trendlines(System.Reflection.Missing.Value)).Add(Type: XlTrendlineType.xlLinear, DisplayEquation: true);//, DisplayRSquared: true);
            return myTrend;
        }

        public void ConfigTrendlineDatalabel(Trendline Trendline)
        {
            Trendline.DataLabel.Top = 10;
            Trendline.DataLabel.Left = 575;
            Trendline.DataLabel.NumberFormat = "0.0000000";
        }

        public double[] GetTrendlineEquationArguments(Microsoft.Office.Interop.Excel.Trendline Trendline)
        {
            string myequation = Trendline.DataLabel.Text;
            string[] equationElements = myequation.Split(' ');
            char xChar = 'x';
            char[] cChar = { '²', 'R', 'n', '\\' };
            string TempString1 = equationElements[2].TrimEnd(xChar);
            string TempString2 = equationElements[4].TrimEnd(cChar);
            double xCoefficient = Convert.ToDouble(TempString1);
            double constant = Convert.ToDouble(TempString2);
            double[] results = new double[2];
            results[0] = xCoefficient;
            results[1] = constant;
            return results;
        }

        public void Save()
        {
            MyWorkbook.Save();
        }

        public void SaveAs(string path)
        {
            MyWorkbook.SaveAs(path);
        }

        public void SelectWorkSheet(int SheetNumber)
        {
            this.MyWorksheet = MyWorkbook.Worksheets[SheetNumber];
        }

        public void DeleteWorkSheet(int SheetNumber)
        {
            MyWorkbook.Worksheets[SheetNumber].Delete();
        }

        public void ProtectSheet()
        {
            MyWorksheet.Protect();
        }

        public void ProtectSheet(string Password)
        {
            MyWorksheet.Protect(Password);
        }

        public void UnprotectSheet()
        {
            MyWorksheet.Unprotect();
        }

        public void UnrotectSheet(string Password)
        {
            MyWorksheet.Unprotect(Password);
        }

        public void Close()
        {
            MyWorkbook.Close();
        }

        public void OpenFile(string filepath, int Sheet)
        {
            ExcelHelper ExcelSheet = new ExcelHelper(filepath, Sheet);
            ExcelSheet.Close();
        }

        public void ReadData(string filepath, int Sheet, int row, int column)
        {
            ExcelHelper ExcelSheet = new ExcelHelper(filepath, Sheet);
            MessageBox.Show(ExcelSheet.ReadCell(row, column));
            ExcelSheet.Close();
        }

        public void WriteDataString(string filepath, int Sheet, int row, int column, string data)
        {
            ExcelHelper ExcelSheet = new ExcelHelper(filepath, Sheet);
            ExcelSheet.WriteToCellString(row, column, data);
            ExcelSheet.Save();
            ExcelSheet.Close();
        }

        public void WriteDataInt(string filepath, int Sheet, int row, int column, int data)
        {
            ExcelHelper ExcelSheet = new ExcelHelper(filepath, Sheet);
            ExcelSheet.WriteToCellInt(row, column, data);
            ExcelSheet.Save();
            ExcelSheet.Close();
        }

        public void WriteDataDouble(string filepath, int Sheet, int row, int column, double data)
        {
            ExcelHelper ExcelSheet = new ExcelHelper(filepath, Sheet);
            ExcelSheet.WriteToCellDouble(row, column, data);
            ExcelSheet.Save();
            ExcelSheet.Close();
        }

        public void CreateNewFile(string filepath)
        {
            ExcelHelper Excel = new ExcelHelper();
            Excel.CreateNewFile();
            Excel.SaveAs(filepath);
            Excel.Close();
        }

        public void CreateNewSheet(string filepath, int Sheet)
        {
            ExcelHelper Excel = new ExcelHelper(filepath, Sheet);
            Excel.CreateNewSheet();
            Excel.Save();
            Excel.Close();
        }

        public string[,] ReadRangeString(string filepath, int Sheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            ExcelHelper ExcelSheet = new ExcelHelper(filepath, Sheet);
            string[,] read = ExcelSheet.ReadRangeString(startRow, startColumn, endRow, endColumn);
            ExcelSheet.Close();
            return read;
        }

        public double[,] ReadRangeDouble(string filepath, int Sheet, int startRow, int startColumn, int endRow, int endColumn)
        {
            ExcelHelper ExcelSheet = new ExcelHelper(filepath, Sheet);
            double[,] read = ExcelSheet.ReadRangeDouble(startRow, startColumn, endRow, endColumn);
            ExcelSheet.Close();
            return read;
        }

        public void CreateExcelFileWithRMS(List<AnalyticsObject> AnalyticsList, string ResultFolderFilePath, double RatedCurrent)
        {
            DateTime CurrentTime = DateTime.Now;
            string TimeFormat = "d_MM_yyyy_HH_mm";
            CreateNewFile($"{ResultFolderFilePath}\\{CurrentTime.ToString(TimeFormat)}");
            ExcelHelper ex1 = new ExcelHelper($"{ResultFolderFilePath}\\{CurrentTime.ToString(TimeFormat)}", 1);

            List<AnalyticsObject> SortedAnalyticsList = AnalyticsList.OrderBy(o => o.EtalonCurrent).ToList();

            int StartLine = 4;
            foreach (AnalyticsObject AnalyticsObject in SortedAnalyticsList)
            {
                ex1.WriteToCellDouble(StartLine, 2, AnalyticsObject.EtalonCurrent); // Действующее значение тока на эталоне

                ex1.WriteToCellDouble(StartLine, 3, 2 * Math.Log(AnalyticsObject.EtalonCurrent * (1 / RatedCurrent), 2)); // Диапазон
                ex1.ChangeCellNumberFormat($"C{StartLine}", "0");

                ex1.WriteToCellDouble(StartLine, 4, AnalyticsObject.Channel1AmplitudeError); // Канал 1 - Амплитудная погрешность
                ex1.WriteToCellDouble(StartLine, 5, AnalyticsObject.Channel2AmplitudeError); // Канал 2 - Амплитудная погрешность
                ex1.WriteToCellDouble(StartLine, 6, AnalyticsObject.Channel3AmplitudeError); // Канал 3 - Амплитудная погрешность
                ex1.WriteToCellDouble(StartLine, 7, AnalyticsObject.Channel4AmplitudeError); // Канал 4 - Амплитудная погрешность

                ex1.WriteToCellDouble(StartLine, 8, 18 * AnalyticsObject.Channel1PhaseError * 1000 * 60); // Канал 1 - Угловая погрешность
                ex1.WriteToCellDouble(StartLine, 9, 18 * AnalyticsObject.Channel2PhaseError * 1000 * 60); // Канал 2 - Угловая погрешность
                ex1.WriteToCellDouble(StartLine, 10, 18 * AnalyticsObject.Channel3PhaseError * 1000 * 60); // Канал 3 - Угловая погрешность
                ex1.WriteToCellDouble(StartLine, 11, 18 * AnalyticsObject.Channel4PhaseError * 1000 * 60); // Канал 4 - Угловая погрешность

                ex1.WriteToCellDouble(StartLine, 13, (1 / (1 - (AnalyticsObject.Channel1AmplitudeError * 0.01)))); // Канал 1 - Поправочный коэффициент
                ex1.WriteToCellDouble(StartLine, 14, (1 / (1 - (AnalyticsObject.Channel2AmplitudeError * 0.01)))); // Канал 2 - Поправочный коэффициент
                ex1.WriteToCellDouble(StartLine, 15, (1 / (1 - (AnalyticsObject.Channel3AmplitudeError * 0.01)))); // Канал 3 - Поправочный коэффициент
                ex1.WriteToCellDouble(StartLine, 16, (1 / (1 - (AnalyticsObject.Channel4AmplitudeError * 0.01)))); // Канал 4 - Поправочный коэффициент

                ex1.WriteToCellDouble(StartLine, 18, AnalyticsObject.Channel1PhaseError * 1000000000); // Канал 1 - Задержка
                ex1.WriteToCellDouble(StartLine, 19, AnalyticsObject.Channel1PhaseError * 1000000000); // Канал 2 - Задержка
                ex1.WriteToCellDouble(StartLine, 20, AnalyticsObject.Channel1PhaseError * 1000000000); // Канал 3 - Задержка
                ex1.WriteToCellDouble(StartLine, 21, AnalyticsObject.Channel1PhaseError * 1000000000); // Канал 4 - Задержка

                StartLine++;
            }

            ex1.ChangeCellNumberFormat("C4", "0");

            ex1.WriteToCellString(2, 2, "Действующее значение тока на эталоне, мкА");
            ex1.WriteToCellString(3, 3, "Диапазон");
            ex1.WriteToCellString(2, 4, "Амплитудная погрешность, %");
            ex1.WriteToCellString(2, 8, "Фазовая погрешность, °мин");
            ex1.WriteToCellString(2, 13, "Поправочный коэффициент");
            ex1.WriteToCellString(2, 18, "Задержка, нс");

            ex1.WriteToCellString(3, 4, "A");
            ex1.WriteToCellString(3, 5, "B");
            ex1.WriteToCellString(3, 6, "C");
            ex1.WriteToCellString(3, 7, "N");

            ex1.WriteToCellString(3, 8, "A");
            ex1.WriteToCellString(3, 9, "B");
            ex1.WriteToCellString(3, 10, "C");
            ex1.WriteToCellString(3, 11, "N");

            ex1.WriteToCellString(3, 13, "A");
            ex1.WriteToCellString(3, 14, "B");
            ex1.WriteToCellString(3, 15, "C");
            ex1.WriteToCellString(3, 16, "N");

            ex1.WriteToCellString(3, 18, "A");
            ex1.WriteToCellString(3, 19, "B");
            ex1.WriteToCellString(3, 20, "C");
            ex1.WriteToCellString(3, 21, "N");

            ex1.MergeCells("D2", "G2");
            ex1.MergeCells("H2", "K2");
            ex1.MergeCells("M2", "P2");
            ex1.MergeCells("R2", "U2");

            ex1.MergeCells("A2", "A3");
            ex1.MergeCells("B2", "B3");
            ex1.MergeCells("C2", "C3");

            int NumberOfPCAPNGFiles = SortedAnalyticsList.Count;

            ex1.AddBordersToRange("B4", $"K{3 + NumberOfPCAPNGFiles}");
            ex1.AddBordersToRange("M4", $"P{3 + NumberOfPCAPNGFiles}");
            ex1.AddBordersToRange("R4", $"U{3 + NumberOfPCAPNGFiles}");

            ex1.AddBoldBordersToRange("B2", "K3");
            ex1.AddBoldBordersToRange("M2", "P3");
            ex1.AddBoldBordersToRange("R2", "U3");

            ex1.AddBoldBordersAroundToRange("B4", $"K{3 + NumberOfPCAPNGFiles}");
            ex1.AddBoldBordersAroundToRange("M4", $"P{3 + NumberOfPCAPNGFiles}");
            ex1.AddBoldBordersAroundToRange("R4", $"U{3 + NumberOfPCAPNGFiles}");


            ex1.AutofitColumns("A1", "V50");
            ex1.CenterAligmentColumns("A1", "V50");

            ex1.WrapTextInCells("A2", "A2");
            ex1.WrapTextInCells("B2", "B2");

            ex1.ChangeColumnWidth(1, 11.5);
            ex1.ChangeColumnWidth(2, 25);
            ex1.ChangeColumnWidth(12, 2.5);
            ex1.ChangeColumnWidth(17, 2.5);

            ex1.Save();
            ex1.Close();
        }
    }
}
