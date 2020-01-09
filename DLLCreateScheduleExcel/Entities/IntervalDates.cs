using System;

namespace DLLCreateScheduleExcel.Entities
{
    class IntervalDates
    {
        public DateTime StartSchedule { get; set; }
        public DateTime EndSchedule { get; set; }

        public IntervalDates(DateTime startSchedule, DateTime endSchedule)
        {
            StartSchedule = startSchedule;
            EndSchedule = endSchedule;
        }

        public IntervalDates()
        {

        }
    }
}
