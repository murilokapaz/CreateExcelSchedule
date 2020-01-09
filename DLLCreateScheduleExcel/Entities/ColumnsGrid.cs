using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DLLCreateScheduleExcel.Entities
{
    class ColumnsGrid
    {
        public int Item { get; set; }
        public string NameActivity { get; set; }
        public string Responsible { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public int WorkingDays { get; set; }
        public int Accomplished { get; set; }
        

        public ColumnsGrid(int item, string nameActivity, string responsible, DateTime startDate, DateTime endDate, int workingDays,int accomplished)
        {
            Item = item;
            NameActivity = nameActivity;
            Responsible = responsible;
            StartDate = startDate;
            EndDate = endDate;
            WorkingDays = workingDays;
            Accomplished = accomplished;
        }

        public ColumnsGrid()
        {

        }
    }
}
