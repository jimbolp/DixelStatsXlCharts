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

                //HumidChartRanges(xlWSheet);
            }
            releaseObject(xlWSheets);

            //SaveAndClose();
        }

        private void HumidChartRanges(Worksheet xlWSheet)
        {
            int startChartPositionLeft = 100;
            int startChartPositionTop = 100;
            ChartObjects xlChartObjs = xlWSheet.ChartObjects();
            ChartRange ChRange = null;
            Range usedRange = xlWSheet.UsedRange;
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
            cInfo.DateTimeFormat.ShortTimePattern = "hh.mm.ss";
            cInfo.DateTimeFormat.DateSeparator = "/";
            DateTime date;
            string currDateCell;
            for (int i = 1; i <= usedRows; ++i)
            {
                MainForm.WriteIntoLabel("Chart " + ChRange.ChartNumber + " ->  Row: " + ChRange.RowOfRange, 1);
                currDateCell = Convert.ToString((usedRange.Cells[i, 1] as Range).Value, cInfo);
                if (currDateCell.Contains("\'"))
                    currDateCell = currDateCell.Remove(currDateCell.IndexOf('\''));
                if (DateTime.TryParse(currDateCell, cInfo, DateTimeStyles.None, out date))
                {
                    //currDateCell = date.ToString("dd/MM/yyyy hh.mm.ss");
                    (usedRange.Cells[i, 1] as Range).ClearFormats();
                    (usedRange.Cells[i, 1] as Range).Value = "\'" + currDateCell;
                    if (isFirstDayOfMonth(currDateCell, cInfo))
                    {
                        if (firstDateOFRange && i != usedRows)
                        {
                            ChRange.ExpandRange(i);
                        }
                        else
                        {
                            Thread chThread = new Thread(() =>
                            {
                                ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                            });
                            chThread.Start();
                            chThread.Join();
                            
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
                            Thread chThread = new Thread(() =>
                            {
                                ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                            });
                            chThread.Start();
                            chThread.Join();
                            startChartPositionTop += 600;
                            ChRange.StartNewRange(i);
                        }
                        else
                        {
                            string nextCell = Convert.ToString((usedRange.Cells[(i + 1), 1] as Range).Value);
                            if (isFirstDayOfMonth(nextCell, cInfo))
                            {
                                Thread chThread = new Thread(() =>
                                {
                                    ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                                });
                                chThread.Start();
                                chThread.Join();
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
                        Thread chThread = new Thread(() =>
                        {
                            ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                        });
                        chThread.Start();
                        chThread.Join();
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
            DateTime start = DateTime.Now;

            int startChartPositionLeft = 100;
            int startChartPositionTop = 100;
            ChartObjects xlChartObjs = xlWSheet.ChartObjects();
            ChartRange ChRange = null;            
            Range usedRange = xlWSheet.UsedRange;
            int time = 1;
            
            start = DateTime.Now;
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
            DateTime date;
            string currDateCell;
            
            for (int i = 1; i <= usedRows; ++i)
            {
                MainForm.WriteIntoLabel(time + "| - 1 -> " + (DateTime.Now - start).ToString(), 2);
                MainForm.WriteIntoLabel("Chart " + ChRange.ChartNumber + " ->  Row: " + ChRange.RowOfRange, 1);      
                currDateCell = Convert.ToString((usedRange.Cells[i, 1] as Range).Value);
                if(currDateCell.Contains("\'"))
                    currDateCell = currDateCell.Remove(currDateCell.IndexOf('\''));
                if (DateTime.TryParse(currDateCell, cInfo, DateTimeStyles.None, out date))
                {
                    (usedRange.Cells[i, 1] as Range).ClearFormats();
                    (usedRange.Cells[i, 1] as Range).Value = "\'" + currDateCell;
                    if (date.DayOfWeek == DayOfWeek.Monday)
                    {
                        if (firstDateOFRange && i != usedRows)
                        {
                            MainForm.WriteIntoLabel(time + "| - 2 -> " + (DateTime.Now - start).ToString(), 2);
                            ChRange.ExpandRange(i);
                        }
                        else
                        {
                            MainForm.WriteIntoLabel(time + "| - 3 -> " + (DateTime.Now - start).ToString(), 2);
                            Thread chThread = new Thread(() =>
                            {
                                ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                            });
                            chThread.Start();
                            chThread.Join();
                            startChartPositionTop += 600;
                            ChRange.StartNewRange(i);
                            firstDateOFRange = true;
                        }
                    }
                    else
                    {
                        MainForm.WriteIntoLabel(time + "| - 4 -> " + (DateTime.Now - start).ToString(), 2);
                        ChRange.ExpandRange(i);
                        firstDateOFRange = false;
                        
                        if (i == usedRows)
                        {
                            MainForm.WriteIntoLabel(time + "| - 5 -> " + (DateTime.Now - start).ToString(), 2);
                            Thread chThread = new Thread(() =>
                            {
                                ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                            });
                            chThread.Start();
                            chThread.Join();
                            startChartPositionTop += 600;
                            ChRange.StartNewRange(i);
                        }
                        else
                        {
                            MainForm.WriteIntoLabel(time + "| - 6 -> " + (DateTime.Now - start).ToString(), 2);
                            DateTime nextDate;
                            string nextCell = Convert.ToString((usedRange.Cells[(i + 1), 1] as Range).Value);                            
                            if (DateTime.TryParse(nextCell, out nextDate) && nextDate.DayOfWeek == DayOfWeek.Monday)
                            {
                                Thread chThread = new Thread(() =>
                                {
                                    ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                                });
                                chThread.Start();
                                chThread.Join();
                                startChartPositionTop += 600;
                                ChRange.StartNewRange(i + 1);
                                firstDateOFRange = true;
                            }
                        }
                    }
                }
                else
                {
                    MainForm.WriteIntoLabel(time + "| - 7 -> " + (DateTime.Now - start).ToString(), 2);
                    if (ChRange.EnoughDataForChart())
                    {
                        Thread chThread = new Thread(() =>
                        {
                            ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                        });
                        chThread.Start();
                        chThread.Join();
                        startChartPositionTop += 600;
                        ChRange.StartNewRange(i);
                        firstDateOFRange = true;
                    }
                    ChRange.StartNewRange(i + 1);
                }
                if (ChRange.ChartNumber > time)
                {
                    time = ChRange.ChartNumber;
                    start = DateTime.Now;
                }
            }
        }
        
        private bool isFirstDayOfMonth(string date, CultureInfo cInfo)
        {
            DateTime d;
            if (DateTime.TryParse(date, cInfo, DateTimeStyles.None, out d) && d.Day == 1)
            {
                return true;
            }
            return false;
        }
        private bool isMonday(string date, CultureInfo cInfo)
        {
            DateTime d;
            if(DateTime.TryParse(date,cInfo, DateTimeStyles.None, out d) && d.DayOfWeek == DayOfWeek.Monday)
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
            string fullSavePath = null;
            try
            {
                MainForm.SaveDialogBox(saveFileDir);
                if(MainForm.SaveFilePath == null)
                {
                    return;
                }
                else
                {
                    try
                    {
                        xlWBook.SaveAs(MainForm.SaveFilePath);
                        MainForm.WriteIntoLabel("File saved successfully in \"" + fullSavePath + "\"", 1);
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
                //MessageBox.Show("The application was already closed or there was a problem closing it!");
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
