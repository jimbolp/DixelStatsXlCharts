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
using System.Threading;
using System.Reflection;

namespace DixelXlCharts
{    
    internal class DixelData
    {
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
                ReleaseObject(xlWBooks);
                xlApp.Quit();
                ReleaseObject(xlApp);
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
            catch(COMException)
            {
                throw;
            }
        }
        public void LoadData()
        {
            List<Thread> treads = new List<Thread>();
            Thread load = new Thread(() =>
            {
                DateTime start = DateTime.Now;
                Sheets xlWSheets = xlWBook.Worksheets;
                foreach (Worksheet xlWSheet in xlWSheets)
                {
                    Thread tr = new Thread(() =>
                    {

                        ConvertDateCellsToText(xlWSheet.UsedRange);
                        TempChartRanges(xlWSheet);

                        HumidChartRanges(xlWSheet);

                    });
                    tr.Start();
                    treads.Add(tr);
                }
                foreach(Thread t in treads)
                {
                    t.Join();
                }
                MainForm.WriteIntoLabel((DateTime.Now - start).ToString(), 2);
            });
            load.Start();
            load.Join();
                           
        }

        private void HumidChartRanges(Worksheet xlWSheet)
        {
            int startChartPositionLeft = 100;
            int startChartPositionTop = 100;
            ChartObjects xlChartObjs = xlWSheet.ChartObjects();
            ChartRange ChRange = null;
            Range usedRange = xlWSheet.UsedRange;
            Range firstCol = usedRange.Columns[1];
            try
            {
                ChRange = new ChartRange('H', usedRange, printNeeded);
            }
            catch (ArgumentException ae)
            {
                MessageBox.Show(ae.Message);
                return;
            }
            int usedRows = usedRange.Rows.Count;
            bool firstDateOFRange = true;
            CultureInfo cInfo = new CultureInfo("bg-BG");
            cInfo.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy";
            cInfo.DateTimeFormat.ShortTimePattern = "hh.mm";
            cInfo.DateTimeFormat.DateSeparator = "/";
            string currDateCell;
            object[,] xlRangeArr = usedRange.Value;
            
            for (int i = 1; i <= usedRows; ++i)
            {
                MainForm.WriteIntoLabel("Chart " + ChRange.ChartNumber + " ->  Row: " + ChRange.RowOfRange, 1);
                currDateCell = Convert.ToString(xlRangeArr[i, 1], cInfo);
                if (currDateCell.Contains("\'"))
                    currDateCell = currDateCell.Remove(currDateCell.IndexOf('\''), 1);
                if (DateTime.TryParse(currDateCell, cInfo, DateTimeStyles.None, out DateTime date))
                {                    
                    if (IsFirstDayOfMonth(currDateCell, cInfo))
                    {
                        if (firstDateOFRange && i != usedRows)
                        {
                            ChRange.ExpandRange(i);
                        }
                        else
                        {
                            ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                            
                            startChartPositionTop += 600;
                            ChRange.StartNewRange(i);
                            firstDateOFRange = true;
                        }
                    }
                    else
                    {
                        ChRange.ExpandRange(i);
                        firstDateOFRange = false;

                        if (i == usedRows)
                        {
                            ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                            
                            startChartPositionTop += 600;
                            ChRange.StartNewRange(i);
                        }
                        else
                        {
                            string nextCell = Convert.ToString(xlRangeArr[i+1, 1]);
                            if (IsFirstDayOfMonth(nextCell, cInfo))
                            {
                                ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                                startChartPositionTop += 600;
                                ChRange.StartNewRange(i + 1);
                                firstDateOFRange = true;
                            }
                        }
                    }
                }
                else
                {
                    if (ChRange.EnoughDataForChart())
                    {
                        ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                        startChartPositionTop += 600;
                        ChRange.StartNewRange(i);
                        firstDateOFRange = true;
                    }
                    ChRange.StartNewRange(i + 1);
                }
            }
        }

        private void TempChartRanges(Worksheet xlWSheet)
        {
            int startChartPositionLeft = 100;
            int startChartPositionTop = 100;
            ChartObjects xlChartObjs = xlWSheet.ChartObjects();
            ChartRange ChRange = null;            
            Range usedRange = xlWSheet.UsedRange;
            Range firstCol = usedRange.Columns[1];

            try
            {
                ChRange = new ChartRange('T', usedRange, printNeeded);
            }
            catch (ArgumentException ae)
            {
                MessageBox.Show(ae.Message);
                return;
            }
            int usedRows = usedRange.Rows.Count;
            bool firstDateOFRange = true;
            CultureInfo cInfo = new CultureInfo("bg-BG");
            cInfo.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy";
            cInfo.DateTimeFormat.ShortTimePattern = "hh.mm.ss";
            cInfo.DateTimeFormat.DateSeparator = "/";
            string currDateCell;
            
            object[,] xlRangeArr = usedRange.Value;
            
            for (int i = 1; i <= usedRows; ++i)
            {
                MainForm.WriteIntoLabel("Chart " + ChRange.ChartNumber + " ->  Row: " + ChRange.RowOfRange, 1);      
                currDateCell = Convert.ToString(xlRangeArr[i,1]);
                if(currDateCell.Contains("\'"))
                    currDateCell = currDateCell.Remove(currDateCell.IndexOf('\''),1);
                if (DateTime.TryParse(currDateCell, cInfo, DateTimeStyles.None, out DateTime date))
                {                    
                    if (date.DayOfWeek == DayOfWeek.Monday)
                    {
                        if (firstDateOFRange && i != usedRows)
                        {
                            ChRange.ExpandRange(i);
                        }
                        else
                        {
                            ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                            startChartPositionTop += 600;
                            ChRange.StartNewRange(i);
                            firstDateOFRange = true;
                        }
                    }
                    else
                    {
                        ChRange.ExpandRange(i);
                        firstDateOFRange = false;
                        
                        if (i == usedRows)
                        {
                            ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                            startChartPositionTop += 600;
                            ChRange.StartNewRange(i);
                        }
                        else
                        {
                            string nextCell = Convert.ToString(xlRangeArr[i+1, 1]);
                            if (DateTime.TryParse(nextCell, out DateTime nextDate) && nextDate.DayOfWeek == DayOfWeek.Monday)
                            {
                                ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                                startChartPositionTop += 600;
                                ChRange.StartNewRange(i + 1);
                                firstDateOFRange = true;
                            }
                        }
                    }
                }
                else
                {
                    if (ChRange.EnoughDataForChart())
                    {
                        ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                        
                        startChartPositionTop += 600;
                        ChRange.StartNewRange(i);
                        firstDateOFRange = true;
                    }
                    ChRange.StartNewRange(i + 1);
                }
            }
        }
        private void ConvertDateCellsToText(Range usedRange)
        {
            object[,] xlNewRange = usedRange.Value;
            for (int i = 1; i <= usedRange.Rows.Count; ++i)
            {
                MainForm.WriteIntoLabel(i.ToString(), 2);
                for (int j = 1; j <= usedRange.Columns.Count; ++j)
                {
                    if (DateTime.TryParse(xlNewRange[i, j].ToString(), out DateTime d))
                        xlNewRange[i, j] = "\'" + xlNewRange[i, j];
                }
            }
            usedRange.Value = xlNewRange;
        }
        private bool IsFirstDayOfMonth(string date, CultureInfo cInfo)
        {
            if (DateTime.TryParse(date, cInfo, DateTimeStyles.None, out DateTime d) && d.Day == 1)
            {
                return true;
            }
            return false;
        }
        private bool IsMonday(string date, CultureInfo cInfo)
        {
            if (DateTime.TryParse(date, cInfo, DateTimeStyles.None, out DateTime d) && d.DayOfWeek == DayOfWeek.Monday)
            {
                return true;
            }
            return false;
        }
        private bool IsSunday(string date)
        {
            if (DateTime.TryParse(date, out DateTime d) && d.DayOfWeek == DayOfWeek.Sunday)
            {
                return true;
            }
            return false;
        }
        public void SaveAndClose()
        {
            try
            {
                MainForm.SaveDialogBox(saveFileDir);
                if(string.IsNullOrEmpty(MainForm.SaveFilePath))
                {
                    MainForm.WriteIntoLabel("File was not saved!", 1);
                    xlWBook.Close(false);
                    xlWBooks.Close();
                    xlApp.Quit();
                    Dispose();
                    return;
                }
                else
                {
                    try
                    {
                        xlWBook.SaveAs(MainForm.SaveFilePath);
                        MainForm.WriteIntoLabel("File saved successfully in \"" + MainForm.SaveFilePath + "\"", 1);
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
                //File probably already closed :D :D
            }
            catch (COMException)
            {
                MessageBox.Show("Unable to close the application or it's already closed! Check Task Manager :D :D");
            }
            ReleaseObject(xlWBook);
            ReleaseObject(xlWBooks);
            ReleaseObject(xlApp);
        }
        private void ReleaseObject(object obj)
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
