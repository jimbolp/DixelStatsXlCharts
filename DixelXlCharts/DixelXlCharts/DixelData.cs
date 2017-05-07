using System;
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
            try
            {
                SetSaveDirectory(filePath);
                printNeeded = print;
                xlWBooks = xlApp.Workbooks;
                LoadFile(filePath);
            }
            catch (ArgumentException)
            {
                MessageBox.Show("Invalid file path!");
                return;
            }
            catch (NullReferenceException)
            {
                ReleaseObject(xlWBooks);
                xlApp.Quit();
                ReleaseObject(xlApp);
                return;
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
                return;
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
            //DateTime start = DateTime.Now;
            Sheets xlWSheets;
                try
                {
                    xlWSheets = xlWBook.Worksheets;
                }
                catch (COMException)
                {
                    throw;
                }
                catch (Exception)
                {
                    throw;
                }

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
                foreach (Thread t in treads)
                {
                    t.Join();
                }
            //MainForm.WriteIntoLabel((DateTime.Now - start).ToString(), 2);
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
            Thread progBar = new Thread(() =>
            {
                MainForm.ProgressBar(usedRows, true);
            });
            progBar.Start();
            progBar.Join();
            bool firstDateOFRange = true;
            CultureInfo cInfo = new CultureInfo("bg-BG");
            cInfo.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy";
            cInfo.DateTimeFormat.ShortTimePattern = "hh.mm";
            cInfo.DateTimeFormat.DateSeparator = "/";
            string currDateCell;
            object[,] xlRangeArr = usedRange.Value;
            
            for (int i = 1; i <= usedRows; ++i)
            {
                MainForm.ProgressBar(i, false);
                //MainForm.WriteIntoLabel("Chart " + ChRange.ChartNumber + " ->  Row: " + ChRange.RowOfRange, 1);
                currDateCell = Convert.ToString(xlRangeArr[i, 1]).Split(new char[0], StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                if (currDateCell.Contains("\'"))
                    currDateCell = currDateCell.Remove(currDateCell.IndexOf('\''), 1);
                DateTime date;
                if (DateTime.TryParse(currDateCell, cInfo, DateTimeStyles.None, out date))
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
            Thread progBar = new Thread(() =>
            {
                MainForm.ProgressBar(usedRows, true);
            });
            progBar.Start();
            progBar.Join();
            bool firstDateOFRange = true;
            CultureInfo cInfo = new CultureInfo("bg-BG");
            cInfo.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy";
            cInfo.DateTimeFormat.ShortTimePattern = "hh.mm.ss";
            cInfo.DateTimeFormat.DateSeparator = "/";
            string currDateCell;
            
            object[,] xlRangeArr = usedRange.Value;
            
            for (int i = 1; i <= usedRows; ++i)
            {
                MainForm.ProgressBar(i, false);
                //MainForm.WriteIntoLabel("Chart " + ChRange.ChartNumber + " ->  Row: " + ChRange.RowOfRange, 1);      
                currDateCell = Convert.ToString(xlRangeArr[i,1]).Split(new char[0], StringSplitOptions.RemoveEmptyEntries)[0].Trim();
                if(currDateCell.Contains("\'"))
                    currDateCell = currDateCell.Remove(currDateCell.IndexOf('\''),1);
                DateTime date;
                if (DateTime.TryParse(currDateCell, cInfo, DateTimeStyles.None, out date))
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
                            DateTime nextDate;
                            if (DateTime.TryParse(nextCell, out nextDate) && nextDate.DayOfWeek == DayOfWeek.Monday)
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
            MainForm.ConvProgBar(0, true);
            MainForm.ConvProgBar(usedRange.Rows.Count, true);

            object[,] xlNewRange = usedRange.Value;
            for (int i = 1; i <= usedRange.Rows.Count; ++i)
            {
                MainForm.ConvProgBar(i, false);
                
                    DateTime d;
                    if (DateTime.TryParse(xlNewRange[i, 1].ToString(), out d))
                        xlNewRange[i, 1] = "\'" + xlNewRange[i, 1];
                
            }
            usedRange.Value = xlNewRange;
            //MainForm.ConvProgBar(0, true);
        }
        private bool IsFirstDayOfMonth(string date, CultureInfo cInfo)
        {
            DateTime d;
            if (DateTime.TryParse(date, cInfo, DateTimeStyles.None, out d) && d.Day == 1)
            {
                return true;
            }
            return false;
        }
        private bool IsMonday(string date, CultureInfo cInfo)
        {
            DateTime d;
            if (DateTime.TryParse(date, cInfo, DateTimeStyles.None, out d) && d.DayOfWeek == DayOfWeek.Monday)
            {
                return true;
            }
            return false;
        }
        private bool IsSunday(string date)
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
            try
            {
                MainForm.SaveDialogBox(saveFileDir);
                if(string.IsNullOrEmpty(MainForm.SaveFilePath))
                {
                    MessageBox.Show("File was not saved!");
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
                        MessageBox.Show("File saved successfully in \"" + MainForm.SaveFilePath + "\"");
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
            catch (Exception)
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
