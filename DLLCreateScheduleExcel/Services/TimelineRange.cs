using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DLLCreateScheduleExcel.Entities;


namespace DLLCreateScheduleExcel.Services
{
    class TimelineRange
    {
        public int ColSpan { get; set; }
        public List<string> TimelineDaysList(IntervalDates interval)
        {
 
            var startDay = interval.StartSchedule;
            TimeSpan diffDays = interval.EndSchedule.Subtract(startDay);
            int numberDays = (int) diffDays.TotalDays;

            List<string> listOfDays = new List<string>();

            var day = startDay;
            listOfDays.Add(day.ToShortDateString());
            for (int i = 0; i<numberDays; i++)
            {
                day = day.AddDays(1);
                listOfDays.Add(day.ToShortDateString());
            }
            return listOfDays;
        }

        public List<string> TimeLineMonthsList(IntervalDates interval)
        {
            string[] month = new string[13] { "", "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro" };


            var startDate = interval.StartSchedule;
            TimeSpan diffDays = interval.EndSchedule.Subtract(startDate);
            int numberDays = (int)diffDays.TotalDays;

            List<string> listOfMonths = new List<string>();

            var day = startDate;
            listOfMonths.Add(month[day.Month]);
            this.ColSpan = 0;
            for (int i = 0; i < numberDays; i++)
            {
                day = day.AddDays(1);
                listOfMonths.Add(month[day.Month]);

            }

            return listOfMonths;

        }

    }
}
