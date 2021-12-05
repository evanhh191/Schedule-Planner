using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace OpenXML_Schedule_project
{
    public partial class Form1 : Form
    {
        /*class Assignment {
            public string whichClass { get; set; }
            public DateTime date { get; set; }
            public string assignment { get; set; }

            public Assignment(string aClass, DateTime aDate, string anAssignemt)
            {
                whichClass = aClass;
                date = aDate;
                assignment = anAssignemt;
            }
        };*/
        private readonly List<Assignment> schedule = new List<Assignment>();      //hidden list of Assignment class, used to store info and to make strings for displayed list(lstAssignments).

        public Form1()
        {
            InitializeComponent();
        }


        private void btnAdd_Click(object sender, EventArgs e)    //adds the info from the class date and time fields to the list and resorts it, *****by date. will add class later*******
        {
            if (cmbClass.Text == "" || txtAssignment.Text == "") //more efficient than creating an object then checking
            {
                MessageBox.Show("Please make sure to fill out the Class and Assignment fields", "Error");
            } else
            {
                schedule.Add(new Assignment(dtpDueDate.Value.Date, cmbClass.Text, txtAssignment.Text));
                if (!cmbClass.Items.Contains(cmbClass.Text))
                {
                    cmbClass.Items.Add(cmbClass.Text);
                }
                schedule.Sort((a, b) => DateTime.Compare(a.Date, b.Date)); // less memory usage than creating another list to sort
                lstAssignmentsBox.Items.Clear();
                foreach (var item in schedule)
                {
                    lstAssignmentsBox.Items.Add(item.Date.ToString("MM/dd/yyyy").PadRight(15) + item.ClassCode.PadRight(26) + item.AssignmentName.PadRight(55)); 
                }
                txtAssignment.Clear();
            }



        }
        /*Assignment newAssignment = new Assignment(cmbClass.Text, dtpDueDate.Value, txtAssignment.Text);

        if (newAssignment.assignment == "" || newAssignment.whichClass == "")
        {
            MessageBox.Show("Please make sure to fill out the Class and Assignment fields", "Error");
        }

        else
        {
            schedule.Add(newAssignment);

            if (!(cmbClass.Items.Contains(cmbClass.Text)))
            {
                cmbClass.Items.Add(cmbClass.Text);
            }

            schedule = schedule.OrderBy(x => x.whichClass).ThenBy(x => x.date).ToList();
            string line;
            lstAssignments.Items.Clear(); 
            foreach (var plan in schedule)
            {
                line = (plan.whichClass.PadRight(26));
                line += plan.date.ToString("MM/dd/yy").PadRight(15);
                line += (plan.assignment.PadRight(55));

                lstAssignments.Items.Add(line);
            }
            txtAssignment.Text = "";
        }
    }*/

        private void btnRemove_Click(object sender, EventArgs e) //Removes selected item from list.
        {
            try
            {
                int selectedIndex = -1;
                selectedIndex = lstAssignmentsBox.SelectedIndex;
                schedule.RemoveAt(selectedIndex);

                schedule.Sort((a, b) => DateTime.Compare(a.Date, b.Date)); // less memory usage than creating another list to sort

                lstAssignmentsBox.Items.Clear();
                foreach (var item in schedule)
                {
                    lstAssignmentsBox.Items.Add(item.Date.ToString("MM/dd/yyyy").PadRight(15) + item.ClassCode.PadRight(26) + item.AssignmentName.PadRight(55));
                }

                //schedule = schedule.OrderBy(x => x.whichClass).ThenBy(x => x.date).ToList();
                /*string line;
                lstAssignmentsBox.Items.Clear();
                foreach (var plan in schedule)
                {
                    line = (plan.whichClass.PadRight(26));
                    line += plan.date.ToString("MM/dd/yy").PadRight(15);
                    line += (plan.assignment.PadRight(55));

                    lstAssignmentsBox.Items.Add(line);
                }*/
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Please select an item on the list to remove.", ex.ToString()); //may as well use ex if we declare it
            }
        }

        private void btnHelp_Click(object sender, EventArgs e)   // Display helpful information for user.
        {
            MessageBox.Show("First you can choose a date using the Due Date button and enter your assigment in the Assignment box. " +
                "\n Next enter the class name in the Class textbox.  You can reselect that class again later after adding an assignment to the list. " +
                "\n When ready, you can click Add to add your assignment to the list. " +
                "\n If you want to remove an item, click on the item in the list and click the remove button. " +
                "\n When you have filled out the list with your assignments, click Build to generate your calender in Excel.", "Schedule Help");
        }

        private void btnBuild_Click(object sender, EventArgs e)  //Once all asisgnments have been entered, this asks the user if they want to make the excel file with the entered information
        {                                                        //If they click yes, then it builds the excel file and exits the program. If no, the dialog closes.
            try
            {

                /*DateTime min = schedule.Min(x => x.Date);
                DateTime max = schedule.Max(x => x.Date);
                int range = max.Subtract(min).Days + 1;*/
                int dateRange = (schedule[^1].Date - schedule[0].Date).Days;


                DialogResult buildResult = MessageBox.Show("Are you ready to create an Excel calendar with the given data?" +
                    "\nYour calendar will start at: " + schedule[0].Date.ToShortDateString() + "\n and end at: " + schedule[^1].Date.ToShortDateString() + "\n For a date range of: " + dateRange, "Build", MessageBoxButtons.YesNo);
                if (buildResult == DialogResult.No)
                {
                }

                else if (buildResult == DialogResult.Yes)       // ** THIS IS WHERE THE SPREADSHEET BUILDING WILL HAPPEN **
                {

                    FolderBrowserDialog browserDialog = new FolderBrowserDialog();
                    string filename;

                    if (browserDialog.ShowDialog() == DialogResult.OK)
                    {
                        filename = browserDialog.SelectedPath;

                        DialogResult locationResult = MessageBox.Show("Save to: " + filename + " ?", "Build", MessageBoxButtons.OKCancel);

                        if (locationResult == DialogResult.OK)
                        {
                            BuildSpreadsheet(filename, dateRange);
                            MessageBox.Show("A spreadsheet calendar has been created at: " + filename);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please ensure that there is at least one entry in the list","Error");
            }
        }

        private void BuildSpreadsheet(string fileName, int dateRange) //builds the spreadsheet
        {
            fileName += "\\SpreadsheetDocumentEx.xlsx";
            try
            {
                IXLWorkbook workbook = new XLWorkbook();
                IXLWorksheet worksheet1 = workbook.Worksheets.Add("Assignments List");

                IXLRange xLRange = worksheet1.Range(worksheet1.Cell(1, 1).Address, worksheet1.Cell(dateRange, 1).Address);
                xLRange.SetDataType(XLDataType.DateTime);


                for (int i = 0; i < schedule.Count; i++)
                {
                    worksheet1.Cell(i + 1, 1).Value = schedule[i].Date.ToString("d");
                    worksheet1.Cell(i + 1, 1).Style.NumberFormat.Format = "m/d/yyyy";
                    worksheet1.Cell(i + 1, 2).Value = schedule[i].ClassCode;
                    worksheet1.Cell(i + 1, 3).Value = schedule[i].AssignmentName;
                }

                //testing formula stuff. will remove
                /*for (int i = schedule.Count + 1; i < 16; i++)
                {
                    string currentCell = "A" + i;
                    worksheet1.Cells(currentCell).FormulaA1 = "=A" + (i - 1) + "+1";
                    worksheet1.Cell(currentCell).Style.NumberFormat.Format = "m/d/yyyy";
                }*/ 

                worksheet1.Columns().AdjustToContents();
                worksheet1.Rows().AdjustToContents();



                workbook.SaveAs(fileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error:", ex.ToString());
            }

        }
    }
}
