using System;

namespace Schedule_Planner
{
    public class Assignment
    {
        private DateTime _date;
        private string _classCode;
        private string _assignmentName;

        public Assignment(DateTime date, string cc, string an)
        {
            this.Date = date;
            this.ClassCode = cc;
            this.AssignmentName = an;
        }

        public DateTime Date
        {
            get
            {
                return this._date;
            }
            set
            {
                this._date = value;
            }
        }

        public string ClassCode
        {
            get
            {
                return this._classCode;
            }
            set
            {
                this._classCode = value;
            }
        }

        public string AssignmentName
        {
            get
            {
                return this._assignmentName;
            }
            set
            {
                this._assignmentName = value;
            }
        }
    }
}