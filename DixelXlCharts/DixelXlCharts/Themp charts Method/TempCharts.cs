private void TempChartRanges(Worksheet xlWSheet)
        {
            //DateTime start = DateTime.Now;

            int startChartPositionLeft = 100;
            int startChartPositionTop = 100;
            ChartObjects xlChartObjs = xlWSheet.ChartObjects();
            ChartRange ChRange = null;            
            Range usedRange = xlWSheet.UsedRange;
            //int time = 1;
            
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
            //start = DateTime.Now;
            for (int i = 1; i <= usedRows; ++i)
            {
                //MainForm.WriteIntoLabel(time + "| - 1 -> " + (DateTime.Now - start).ToString(), 2);
                MainForm.WriteIntoLabel("Chart " + ChRange.ChartNumber + " ->  Row: " + ChRange.RowOfRange, 1);      
                currDateCell = Convert.ToString((usedRange.Cells[i, 1] as Range).Value);
                if(currDateCell.Contains("\'"))
                    currDateCell = currDateCell.Remove(currDateCell.IndexOf('\''));
                if (DateTime.TryParse(currDateCell, cInfo, DateTimeStyles.None, out DateTime date))
                {
                    //(usedRange.Cells[i, 1] as Range).ClearFormats();
                    (usedRange.Cells[i, 1] as Range).Value = "\'" + currDateCell;
                    if (date.DayOfWeek == DayOfWeek.Monday)
                    {
                        if (firstDateOFRange && i != usedRows)
                        {
                            //MainForm.WriteIntoLabel(time + "| - 2 -> " + (DateTime.Now - start).ToString(), 2);
                            ChRange.ExpandRange(i);
                        }
                        else
                        {
                            //MainForm.WriteIntoLabel(time + "| - 3 -> " + (DateTime.Now - start).ToString(), 2);
                            ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                            startChartPositionTop += 600;
                            ChRange.StartNewRange(i);
                            firstDateOFRange = true;
                        }
                    }
                    else
                    {
                        //MainForm.WriteIntoLabel(time + "| - 4 -> " + (DateTime.Now - start).ToString(), 2);
                        ChRange.ExpandRange(i);
                        firstDateOFRange = false;
                        
                        if (i == usedRows)
                        {
                            //MainForm.WriteIntoLabel(time + "| - 5 -> " + (DateTime.Now - start).ToString(), 2);
                            ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                            startChartPositionTop += 600;
                            ChRange.StartNewRange(i);
                        }
                        else
                        {
                            //MainForm.WriteIntoLabel(time + "| - 6 -> " + (DateTime.Now - start).ToString(), 2);
                            string nextCell = Convert.ToString((usedRange.Cells[(i + 1), 1] as Range).Value);
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
                    //MainForm.WriteIntoLabel(time + "| - 7 -> " + (DateTime.Now - start).ToString(), 2);
                    if (ChRange.EnoughDataForChart())
                    {
                        ChRange.CreateChart(xlChartObjs, xlWSheet.Name, startChartPositionLeft, startChartPositionTop);
                        
                        startChartPositionTop += 600;
                        ChRange.StartNewRange(i);
                        firstDateOFRange = true;
                    }
                    ChRange.StartNewRange(i + 1);
                }
                /*
                if (ChRange.ChartNumber > time)
                {
                    time = ChRange.ChartNumber;
                    start = DateTime.Now;
                }//*/
            }
        }