using System;

namespace Schedule_Planner
{
    public class Assignment
    {
        public DateTime StartDate { get; set; }
        public DateTime Date { get; set; }

        public string ClassCode { get; set; }
        public string AssignmentName { get; set; }

        public Assignment(DateTime date, string cc, string an)
        {
            this.Date = date;
            this.ClassCode = cc;
            this.AssignmentName = an;
        }

        public Assignment(DateTime startDate, DateTime date, string cc, string an)
        {
            this.StartDate = startDate;
            this.Date = date;
            this.ClassCode = cc;
            this.AssignmentName = an;
        }
    }
}