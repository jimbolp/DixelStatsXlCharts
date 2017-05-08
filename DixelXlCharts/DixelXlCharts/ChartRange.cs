using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace DixelXlCharts
{
    class ChartRange
    {
        const double chartHeigth = 521.0134; //18.23cm * 28.58
        const double chartWidth = 867.9746; //30.37cm * 28.58
        string topDateCell;
        string topDataCell;
        string bottomDateCell;
        string bottomDataCell;
        Range usedRange;
        Range DateRange;
        Range DataRange;
        readonly bool printNeeded;
        readonly char col;
        readonly int type = 0;
        public int ChartNumber { get; set; } = 1;
        public int RowOfRange { get; set; } = 0;
        public ChartRange(char type, Range usedRange, bool print, bool special)
        {
            if(!(type.ToString().ToLower() == "t" || type.ToString().ToLower() == "h"))
            {
                throw new ArgumentException("The given graph type is not supported! The supported graph types are 'T' for Temperatures or 'H' for Humidity.");
            }
            switch (type.ToString().ToLower())
            {
                case "t":
                    col = 'B';
                    break;
                case "h":
                    {
                        col = special ? 'B' : 'C';
                    }
                    break;
            }
            //col = type.ToString().ToLower() == "t" ? 'B' : 'C';
            this.type = type.ToString().ToLower() == "t" ? 1 : 2;
            topDateCell = "A" + 1;
            topDataCell = col.ToString() + 1;
            bottomDateCell = topDateCell;
            bottomDataCell = topDataCell;
            this.usedRange = usedRange;
            printNeeded = print;
            DateRange = usedRange.Range[topDateCell, bottomDateCell];
            DataRange = usedRange.Range[topDataCell, bottomDataCell];
        }
        public void ExpandRange(int row)
        {
            RowOfRange++;
            bottomDateCell = "A" + row;
            bottomDataCell = col.ToString() + row;
        }
        public void StartNewRange(int row)
        {
            RowOfRange = 1;
            topDateCell = "A" + row;
            topDataCell = col.ToString() + row;
            bottomDateCell = "A" + row;
            bottomDataCell = col.ToString() + row;
        }
        public void CreateChart(ChartObjects xlChartObjs, string Name, double startChartPositionLeft, double startChartPositionTop)
        {
            ChartNumber++;
            if (type == 2) { startChartPositionLeft += 100; } else { startChartPositionTop += 50; } 
            Name = type != 0 ? (type == 1 ? Name + "_T" : Name + "_H") : Name;
            DateRange = usedRange.Range[topDateCell, bottomDateCell];
            DataRange = usedRange.Range[topDataCell, bottomDataCell];
            ChartObject xlChartObj = xlChartObjs.Add(startChartPositionLeft, startChartPositionTop, chartWidth, chartHeigth);
            Chart xlChartPage = xlChartObj.Chart;
            Series xlChartSeries = xlChartPage.SeriesCollection().Add(DataRange);
            xlChartSeries.XValues = DateRange;
            xlChartPage.ChartType = XlChartType.xlLine;
            xlChartPage.HasTitle = true;
            xlChartPage.ChartTitle.Text = Name;
            xlChartPage.Legend.Delete();
            
            if (printNeeded)
                xlChartPage.PrintOut();

            DateRange = null;
            DataRange = null;
            Marshal.ReleaseComObject(xlChartObj);
            xlChartObj = null;
        }
        public bool EnoughDataForChart()
        {
            if (topDateCell == bottomDateCell || topDataCell == bottomDataCell)
            {
                return false;
            }
            return true;
        }
    }
}
