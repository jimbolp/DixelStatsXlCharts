using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace DixelXlCharts
{
    class ChartRange
    {
        string topDateCell;
        string topDataCell;
        string bottomDateCell;
        string bottomDataCell;
        readonly char col;
        public ChartRange(char type)
        {
            if(!(type.ToString().ToLower() == "t" || type.ToString().ToLower() == "h"))
            {
                throw new ArgumentException("The given graph type is not supported! The supported graph types are 'T' for Temperatures or 'H' for Humidity.");
            }
            col = type.ToString().ToLower() == "t" ? 'B' : 'C';
            topDateCell = "A" + 1;
            topDataCell = col.ToString() + 1;
            bottomDateCell = topDateCell;
            bottomDataCell = topDataCell;

        }

        public string TopDate {
            get
            {
                return topDateCell;
            }
            set
            {
                topDateCell = "A" + value;
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
    }
}
