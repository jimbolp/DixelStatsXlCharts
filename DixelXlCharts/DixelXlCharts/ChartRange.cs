using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
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
        public ChartRange(char type, Range usedRange, bool print)
        {
            if(!(type.ToString().ToLower() == "t" || type.ToString().ToLower() == "h"))
            {
                throw new ArgumentException("The given graph type is not supported! The supported graph types are 'T' for Temperatures or 'H' for Humidity.");
            }
            col = type.ToString().ToLower() == "t" ? 'B' : 'C';
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

        public object TopDate {
            get
            {
                return topDateCell;
            }
            set
            {
                topDateCell = "A" + value.ToString();
            }
        }
        public object TopData
        {
            get
            {
                return topDataCell;
            }
            set
            {
                topDataCell = col + value.ToString();
            }
        }
        public object BottomDate
        {
            get
            {
                return bottomDateCell;
            }
            set
            {
                bottomDateCell = "A" + value.ToString();
            }
        }
        public object BottomData
        {
            get
            {
                return bottomDataCell;
            }
            set
            {
                bottomDataCell = col + value.ToString();
            }
        }
        public void ExpandRange(int row)
        {
            RowOfRange++;
            BottomDate = row;
            BottomData = row;
        }
        public void StartNewRange(int row)
        {
            RowOfRange = 1;
            TopDate = row;
            TopData = row;
            BottomDate = row;
            BottomData = row;
        }
        public void CreateChart(ChartObjects xlChartObjs, string Name, double startChartPositionLeft, double startChartPositionTop)
        {
            ChartNumber++;
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
