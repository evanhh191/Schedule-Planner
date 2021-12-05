using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace OpenXML_Schedule_project
{
    public partial class Form1 : Form
    {
        class Assignment {
            public string whichClass { get; set; }
            public DateTime date { get; set; }
            public string assignment { get; set; }

            public Assignment(string aClass, DateTime aDate, string anAssignemt)
            {
                whichClass = aClass;
                date = aDate;
                assignment = anAssignemt;
            }
        };

        public Form1()
        {
            InitializeComponent();
        }

        List<Assignment> schedule = new List<Assignment>();      //hidden list of Assignment class, used to store info and to make strings for displayed list(lstAssignments).

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnAdd_Click(object sender, EventArgs e)    //adds the info from the class date and time fields to the list and resorts it, first by class then by date.
        {
            Assignment newAssignment = new Assignment(cmbClass.Text, dtpDueDate.Value, txtAssignment.Text);

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
        }

        private void btnRemove_Click(object sender, EventArgs e) //Removes selected item from list.
        {
            try
            {
                int selectedIndex = -1;
                selectedIndex = lstAssignments.SelectedIndex;
                schedule.RemoveAt(selectedIndex);

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
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please select an item on the list to remove.", "Error");
            }
        }

        private void btnHelp_Click(object sender, EventArgs e)   // Display helpful information for user.
        {
            MessageBox.Show("First enter the class name in the Class textbox. You can reselect that class again later after adding an assignment to the list. " +
                "\n You can choose a date using the Due Date button and enter your assigment in the Assignment box. " +
                "\n When ready, you can click Add to add your assignment to the list. " +
                "\n If you want to remove an item, click on the item in the list and click the remove button. " +
                "\n When you have filled out the list with your assignments, click Build to generate your calender in Excel.", "Schedule Help");
        }

        private void btnBuild_Click(object sender, EventArgs e)  //Once all asisgnments have been entered, this asks the user if they want to make the excel file with the entered information
        {                                                        //If they click yes, then it builds the excel file and exits the program. If no, the dialog closes.
            try
            {

                DateTime min = schedule.Min(x => x.date);
                DateTime max = schedule.Max(x => x.date);
                int range = max.Subtract(min).Days + 1;

                DialogResult buildResult = MessageBox.Show("Are you ready to create an Excel calendar with the given data?" +
                    "\nYour calendar will start at: " + min.ToShortDateString() + "\n and end at: " + max.ToShortDateString() + "\n For a date range of: " + range, "Build", MessageBoxButtons.YesNo);
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
                            buildSpreadsheet(filename, range);
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

        private void buildSpreadsheet(string fileName, int range) //builds the spreadsheet
        {
            fileName = fileName + "\\SpreadsheetDocumentEx.xlsx";

            DateTime date = dtpDueDate.Value;
            MessageBox.Show(date.ToString());


            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.
               Create(fileName, SpreadsheetDocumentType.Workbook);
            // Add a WorkbookPart to the document.  
            WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart.  
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.  
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.  
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "mySheet"
            };
            sheets.Append(sheet);

            Worksheet worksheet = new Worksheet();
            SheetData sheetData = new SheetData();
            Row row = new Row();
            Cell cell = new Cell()
            {
                CellReference = "A1",
                DataType = CellValues.Date,
                CellValue = new CellValue(date)
            };
            row.Append(cell);
            sheetData.Append(row);
            worksheet.Append(sheetData);
            worksheetPart.Worksheet = worksheet;

            spreadsheetDocument.Close();
        }
    }
}
