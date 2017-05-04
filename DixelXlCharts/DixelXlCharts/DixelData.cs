﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Globalization;

namespace DixelXlCharts
{    
    internal class DixelData
    {
        const double chartHeigth = 521.0134; //18.23cm * 28.58
        const double chartWidth = 867.9746; //30.37cm * 28.58
        readonly bool printNeeded = false;
        string saveFileDir = null;
        Application xlApp = new Application();
        Workbooks xlWBooks = null;
        Workbook xlWBook = null;
        
        public DixelData(string filePath, bool print)
        {
            SetSaveDirectory(filePath);
            printNeeded = print;
            try
            {
                xlWBooks = xlApp.Workbooks;
                LoadFile(filePath);
            }
            catch (NullReferenceException)
            {
                releaseObject(xlWBooks);
                xlApp.Quit();
                releaseObject(xlApp);
            }
            catch (COMException)
            {
                throw new Exception("Invalid File Path.. Object was not created..");
            }
        }
        private void SetSaveDirectory(string path)
        {
            try
            {
                saveFileDir = Path.GetDirectoryName(path);
            }
            catch (ArgumentException)
            {
                MessageBox.Show("Invalid file path!");
            }
            catch (PathTooLongException)
            {
                MessageBox.Show("File path too long!");
            }
        }
        private void LoadFile(string filePath)
        {
            try
            {
                xlWBook = xlWBooks.Open(filePath, IgnoreReadOnlyRecommended: true, ReadOnly: false, Editable: true);
            }
            catch(COMException ex)
            {
                MessageBox.Show("File could not load! " + ex.ToString());
                throw ex;
            }
        }
        public void LoadData()
        {
            Sheets xlWSheets = xlWBook.Worksheets;
            foreach(Worksheet xlWSheet in xlWSheets)
            {
                TempChartRanges(xlWSheet);

                HumidChartRanges(xlWSheet);
            }
            releaseObject(xlWSheets);
        }

        private void HumidChartRanges(Worksheet xlWsheet)
        {
            
        }

        private void TempChartRanges(Worksheet xlWSheet)
        {
            int startChartPositionLeft = 100;
            int startChartPositionTop = 100;
            ChartObjects xlChartObjs = xlWSheet.ChartObjects();

            ChartRange ChRange = new ChartRange('T');
            Range usedRange = xlWSheet.UsedRange;
            Range DateRange = usedRange.Range[topDateCell, bottomDateCell];
            Range DataRange = usedRange.Range[topDataCell, bottomDataCell];
            int usedRows = usedRange.Rows.Count;
            bool firstDateOFRange = true;
            for(int i = 2; i <= usedRows; ++i)
            {
                DateTime date;
                string currDateCell = Convert.ToString((usedRange.Cells[i, 1] as Range).Value);
                if (DateTime.TryParse(currDateCell, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                {
                    currDateCell = date.ToString("dd/MM/yyyy hh.mm.ss");
                    (usedRange.Cells[i, 1] as Range).ClearFormats();
                    (usedRange.Cells[i, 1] as Range).Value = currDateCell;
                    if (isMonday(currDateCell))
                    {
                        if (firstDateOFRange && i != usedRows)
                        {
                            bottomDateCell = "A" + i;
                            bottomDataCell = "B" + i;
                        }
                        else
                        {
                            DateRange = usedRange.Range[topDateCell, bottomDateCell];
                            DataRange = usedRange.Range[topDataCell, bottomDataCell];
                            ChartObject xlChartObj = xlChartObjs.Add(startChartPositionLeft, startChartPositionTop, chartWidth, chartHeigth);
                            Series xlChartSeries = xlChartObj.Chart.SeriesCollection().Add(DataRange);
                            xlChartSeries.XValues = DateRange;
                            CreateChartFromRange(DateRange, DataRange, xlChartObj, xlWSheet.Name + "_T");
                            startChartPositionTop += 600;
                            topDateCell = "A" + i;
                            topDataCell = "B" + i;
                            bottomDateCell = topDateCell;
                            bottomDataCell = topDataCell;
                            firstDateOFRange = true;
                        }
                    }
                    else
                    {
                        bottomDateCell = "A" + i;
                        bottomDataCell = "B" + i;
                        firstDateOFRange = false;
                        string nextCell;
                        if (i == usedRows)
                        {
                            DateRange = usedRange.Range[topDateCell, bottomDateCell];
                            DataRange = usedRange.Range[topDataCell, bottomDataCell];
                            ChartObject xlChartObj = xlChartObjs.Add(startChartPositionLeft, startChartPositionTop, chartWidth, chartHeigth);
                            Series xlChartSeries = xlChartObj.Chart.SeriesCollection().Add(DataRange);
                            xlChartSeries.XValues = DateRange;
                            CreateChartFromRange(DateRange, DataRange, xlChartObj, xlWSheet.Name + "_T");
                            startChartPositionTop += 600;
                            topDateCell = "A" + i;
                            topDataCell = "B" + i;
                            bottomDateCell = topDateCell;
                            bottomDataCell = topDataCell;
                        }
                        else
                        {
                            nextCell = Convert.ToString((usedRange.Cells[(i + 1), 1] as Range).Value);
                            if (isMonday(nextCell))
                            {
                                DateRange = usedRange.Range[topDateCell, bottomDateCell];
                                DataRange = usedRange.Range[topDataCell, bottomDataCell];
                                ChartObject xlChartObj = xlChartObjs.Add(startChartPositionLeft, startChartPositionTop, chartWidth, chartHeigth);
                                Series xlChartSeries = xlChartObj.Chart.SeriesCollection().Add(DataRange);
                                xlChartSeries.XValues = DateRange;
                                CreateChartFromRange(DateRange, DataRange, xlChartObj, xlWSheet.Name + "_T");
                                startChartPositionTop += 600;
                                topDateCell = "A" + i;
                                topDataCell = "B" + i;
                                bottomDateCell = topDateCell;
                                bottomDataCell = topDataCell;
                                firstDateOFRange = true;
                            }
                        }
                    }
                }
                else
                {
                    DateRange = usedRange.Range[topDateCell, bottomDateCell];
                    DataRange = usedRange.Range[topDataCell, bottomDataCell];
                    ChartObject xlChartObj = xlChartObjs.Add(startChartPositionLeft, startChartPositionTop, chartWidth, chartHeigth);
                    Series xlChartSeries = xlChartObj.Chart.SeriesCollection().Add(DataRange);
                    xlChartSeries.XValues = DateRange;
                    CreateChartFromRange(DateRange, DataRange, xlChartObj, xlWSheet.Name + "_T");
                    startChartPositionTop += 600;
                    topDateCell = "A" + i;
                    topDataCell = "B" + i;
                    bottomDateCell = topDateCell;
                    bottomDataCell = topDataCell;
                    firstDateOFRange = true;
                }
            }
        }
        
        private void CreateChartFromRange(Range X_Series, Range Y_Series, ChartObject xlChartObj, string sheetName)
        {

            xlChartObj.Chart.ChartType = XlChartType.xlLine;
            xlChartObj.Chart.HasTitle = true;
            Chart xlChartPage = xlChartObj.Chart;
            xlChartPage.ChartTitle.Text = sheetName;
            xlChartPage.Legend.Delete();

            if (printNeeded)
                xlChartPage.PrintOut();
        }
        private bool isMonday(string date)
        {
            DateTime d;
            if(DateTime.TryParse(date, out d) && d.DayOfWeek == DayOfWeek.Monday)
            {
                return true;
            }
            return false;
        }
        private bool isSunday(string date)
        {
            DateTime d;
            if (DateTime.TryParse(date, out d) && d.DayOfWeek == DayOfWeek.Sunday)
            {
                return true;
            }
            return false;
        }
        public void SaveAndClose()
        {
            bool save = false;
            string fullSavePath = null;
            try
            {
                while (!save)
                {
                    var saveFileDialog = new SaveFileDialog
                    {
                        Filter = "Excel Workbook|*.xlsx; *.xlsm|Excel 97-2003 Workbook|*.xls",
                        Title = "Save As",
                        DefaultExt = "xlsx",
                        InitialDirectory = saveFileDir
                    };
                    saveFileDialog.AddExtension = true;
                    
                    DialogResult dr = saveFileDialog.ShowDialog();
                    if (dr == DialogResult.OK)
                    {
                        if (!File.Exists(saveFileDialog.FileName))
                        {
                            save = true;
                            fullSavePath = saveFileDialog.FileName;
                        }
                    }
                    else
                    {
                        try
                        {
                            xlWBook.SaveAs(fullSavePath);
                            xlWBook.Close(false);
                            xlWBooks.Close();
                            xlApp.Quit();
                        }
                        catch (COMException)
                        {
                            Dispose();
                        }
                        return;
                    }
                }
                if (!string.IsNullOrEmpty(fullSavePath))
                {
                    try
                    {
                        xlWBook.SaveAs(fullSavePath);
                        xlWBook.Close(false);
                        xlWBooks.Close();
                        xlApp.Quit();
                    }
                    catch (COMException)
                    {
                        Dispose();
                    }
                }
            }
            catch (COMException comEx)
            {
                MessageBox.Show("An exception was thrown while saving the file:" +
                    Environment.NewLine +
                    comEx.ToString());
            }
            catch (Exception e)
            {
                MessageBox.Show("An exception was thrown while saving the file:" +
                    Environment.NewLine +
                    e.ToString());
            }
        }
        public void Dispose()
        {
            try
            {
                xlApp.Quit();
            }
            catch (InvalidComObjectException)
            {
                MessageBox.Show("The application was already closed or there was a problem closing it!");
            }
            catch (COMException)
            {
                MessageBox.Show("Unable to close the application!");
            }
            releaseObject(xlWBook);
            releaseObject(xlWBooks);
            releaseObject(xlApp);
        }
        private void releaseObject(object obj)
        {
            try
            {
                Marshal.FinalReleaseComObject(obj);
                obj = null;
            }
            catch (COMException cEx)
            {
                obj = null;
                MessageBox.Show("Com Exception Occured while releasing object " + cEx.ToString());
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }
    }
}
