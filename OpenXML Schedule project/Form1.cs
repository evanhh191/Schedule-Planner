using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace OpenXML_Schedule_project
{
    public partial class Form1 : Form
    {
        private readonly List<Assignment> schedule = new List<Assignment>();      //hidden list of Assignment class, used to store info and to make strings for displayed list(lstAssignments).

        public Form1()
        {
            InitializeComponent();
        }

        private void BtnAdd_Click(object sender, EventArgs e)    //adds the info from the date, class, and time fields to the list and resorts it by date then classCode
        {
            if (cmbClass.Text == "" || txtAssignment.Text == "") //more efficient than creating an object then checking
            {
                MessageBox.Show("Please make sure to fill out the Class and Assignment fields", "Error");
            }
            else
            {
                schedule.Add(new Assignment(dtpDueDate.Value.Date, cmbClass.Text, txtAssignment.Text));
                if (!cmbClass.Items.Contains(cmbClass.Text)) cmbClass.Items.Add(cmbClass.Text);

                PrintToList();
                txtAssignment.Clear();
                txtAssignment.Focus();
            }
        }

        private void BtnRemove_Click(object sender, EventArgs e) //Removes selected item from list.
        {
            try
            {
                int selectedIndex = -1;
                selectedIndex = lstAssignmentsBox.SelectedIndex;
                schedule.RemoveAt(selectedIndex);

                PrintToList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please select an item on the list to remove.", "Error");
                Console.WriteLine(ex.ToString()); //may as well use ex if we declare it
            }
        }

        private void BtnHelp_Click(object sender, EventArgs e)   // Display helpful information for user.
        {
            MessageBox.Show("First choose a date using the Due Date button and enter the assignment name in the Assignment box. " +
                "\n\nNext enter the class name in the Class box. You can reselect that class again later after adding an assignment to the list. " +
                "\n\nAlternatively, if you have a pre-existing text or Excel file that is formatted correctly you may use the menu in the upper left hand corner to" +
                " upload data into the list. " +
                "\n\nWhen ready, you can click Add to add your assignment to the list or press the Enter key after you finish typing in the assignment name. " +
                "\n\nIf you want to remove an item, click on the assignment in the list and click the Remove button. " +
                "\n\nWhen you have filled out the list with your assignments, click Build to choose a file location for the program to generate the calendar-containing" +
                " Excel file. ", "Schedule Help");
        }

        private void BtnBuild_Click(object sender, EventArgs e)  //Once all asisgnments have been entered, this asks the user if they want to make the excel file with the entered information
        {                                                        //If they click yes, then it builds the excel file and exits the program. If no, the dialog closes.
            try
            {
                //MessageBox.Show("day of the week for start day:" + (int)schedule[0].Date.DayOfWeek); //idle testing some things
                int dateRange = (int)(schedule[^1].Date.ToOADate() - schedule[0].Date.ToOADate() + 1); //converting to serialized date. Otherwise doesn't work. adding one to be inclusive of start date

                DialogResult buildResult = MessageBox.Show("Are you ready to create an Excel calendar with the given data?" +
                    "\nYour calendar will range from: " + schedule[0].Date.ToShortDateString() + "to: " + schedule[^1].Date.ToShortDateString()
                         + "\nFor a date range of: " + dateRange + " day(s)" + "\nNumber of assignments: " + schedule.Count, "Build", MessageBoxButtons.YesNo);
                if (buildResult == DialogResult.No) { }
                else if (buildResult == DialogResult.Yes)       // ** THIS IS WHERE THE SPREADSHEET BUILDING WILL HAPPEN **
                {
                    FolderBrowserDialog browserDialog = new FolderBrowserDialog();
                    string filename;

                    if (browserDialog.ShowDialog() == DialogResult.OK)
                    {
                        filename = browserDialog.SelectedPath;

                        DialogResult locationResult = MessageBox.Show("Please close spreadsheet if open. \nSave to: " + filename + " ?", "Build", MessageBoxButtons.OKCancel);

                        if (locationResult == DialogResult.OK)
                        {
                            BuildSpreadsheet(filename, dateRange);
                            MessageBox.Show("A spreadsheet calendar has been created at: " + filename); //still pops up even if BuildSpreadsheet(...) catches an error when trying to save spreadsheet
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Please ensure that there is at least one entry in the list", "Error");
                Console.WriteLine(ex.ToString());
            }
        }

        private void BuildSpreadsheet(string fileName, int dateRange) //builds the spreadsheet
        {
            fileName += "\\Schedule.xlsx";
            try
            {
                IXLWorkbook workbook = new XLWorkbook();
                IXLWorksheet worksheet1 = workbook.Worksheets.Add("Sheet1");
                IXLWorksheet worksheet2 = workbook.Worksheets.Add("Sheet2");

                //prepping for data entry
                worksheet1.Column(1).SetDataType(XLDataType.DateTime);

                //styling
                IXLRange headerRange1 = worksheet1.Range(worksheet1.Cell(1, 1).Address, worksheet1.Cell(1, 3).Address);
                headerRange1.Cells().Style.Fill.SetBackgroundColor(XLColor.LightGray);
                headerRange1.Cells().Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                headerRange1.Cells().Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                headerRange1.Cells().Style.Border.OutsideBorderColor = XLColor.Black;
                headerRange1.Cells().Style.Border.InsideBorderColor = XLColor.Black;

                //changing header to text datatype
                worksheet1.Cell(1, 1).SetDataType(XLDataType.Text);
                worksheet1.Cell("A1").Value = "Due";
                worksheet1.Cell("B1").Value = "Class";
                worksheet1.Cell("c1").Value = "Work";

                //filling cells with assignments
                for (int i = 0; i < schedule.Count; i++)
                {
                    worksheet1.Cell(i + 2, 1).Value = schedule[i].Date.ToString("d");
                    worksheet1.Cell(i + 2, 1).Style.NumberFormat.Format = "d-mmm";
                    worksheet1.Cell(i + 2, 2).Value = schedule[i].ClassCode;
                    worksheet1.Cell(i + 2, 3).Value = schedule[i].AssignmentName;
                }

                // Add filters
                worksheet1.RangeUsed().SetAutoFilter();
                // Sort the filtered list
                worksheet1.AutoFilter.Sort(1);

                //day of week header
                for (int i = 0; i < 7; i++)
                {
                    worksheet2.Range(worksheet2.Cell(1, 2 * i + 1), worksheet2.Cell(1, 2 * i + 2)).Merge().Style
                        .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                        .Fill.SetBackgroundColor(XLColor.GoldenYellow)
                        .Border.SetOutsideBorder(XLBorderStyleValues.Thick);
                    worksheet2.Cell(1, 2 * i + 1).Value = (DayOfWeek)i;
                }
                //Boolean previous = false;
                //int previousDateRange = 0;
                //int days = 0;
                int max = dateRange;
                //if (previous) days = previousDateRange;

                max += (int)schedule[0].Date.DayOfWeek;
                var startCell = worksheet2.Cell(2 + ((int)schedule[0].Date.DayOfWeek / 7) * 7, 2 * (((int)schedule[0].Date.DayOfWeek % 7) + 1));
                startCell.Value = schedule[0].Date.ToString("d");
                Boolean first = true;
                int dow = (int)schedule[0].Date.DayOfWeek;

                for (int i = dow; i < max; i++)
                {
                    //(i/7) * 7 counts the number of weeks so far. ex day 6, which is a saturday, is week 0 because for ints, 6/7 = 0
                    int rowIncrementer = (i / 7) * 7;
                    var cellRange = worksheet2.Range(worksheet2.Cell(2 + rowIncrementer, 2 * (i % 7) + 1), worksheet2.Cell(2 + rowIncrementer, 2 * (i % 7) + 2));
                    cellRange.Merge().Style
                        .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                        .Fill.SetBackgroundColor(XLColor.LightGreen)
                        .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                        .NumberFormat.SetFormat("m/d");
                    if (first)
                    {
                        worksheet2.Cell(2 + rowIncrementer, 2 * (i % 7) + 1).Value = schedule[i - dow].Date.ToShortDateString();
                        first = false;
                    }
                    else if (i / 7 < 1 || i / 7 == 1 && i % 7 != 0)
                    {
                        worksheet2.Cell(2 + rowIncrementer, 2 * (i % 7) + 1).FormulaR1C1 = "=RC[-2]+1";
                    }
                    else if (i / 7 == 1 && i % 7 == 0)
                    {
                        worksheet2.Cell(2 + rowIncrementer, 2 * (i % 7) + 1).FormulaR1C1 = "=R[-7]C[12] + 1";
                    }
                    else
                    {
                        worksheet2.Cell(2 + rowIncrementer, 2 * (i % 7) + 1).FormulaR1C1 = "=R[-7]C+7";
                    }
                    worksheet2.Cell(3 + rowIncrementer, 2 * (i % 7) + 1).FormulaR1C1 = "COUNTIF(Sheet1!R2C1:R" + (dateRange + 1) + "C1,R[-1]C)";
                    //worksheet2.Cell(3 + rowIncrementer, 2 * (i % 7) + 2).FormulaR1C1 = "=Sheet1!R[-1]C[-3]";
                }

                worksheet1.Columns().AdjustToContents();
                worksheet1.Rows().AdjustToContents();
                worksheet2.Columns().AdjustToContents();
                worksheet2.Rows().AdjustToContents();

                worksheet1.SheetView.FreezeRows(1);
                worksheet2.SheetView.FreezeRows(1);
                worksheet2.RecalculateAllFormulas();
                workbook.SaveAs(fileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Build failed. File is likely open, but see console logs for details", "Error");
                Console.WriteLine(ex.ToString());
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 160; i++)
            {
                schedule.Add(new Assignment(DateTime.Now.AddDays(i%40), "Test" + i % 4, "test assignment name" + i));
            }
            for (int i = 160; i < 234; i++)
            {
                schedule.Add(new Assignment(DateTime.Now.AddDays(i), "Test" + i % 4, "test assignment name" + i));
            }
            PrintToList();
        }

        private void TxtAssignment_KeyPress(object sender, KeyPressEventArgs e) // If txtAssignment is focus, pressing Enter will attempt to add current info to list
        {
            if (e.KeyChar == '\r')
            {
                if (cmbClass.Text == "" || txtAssignment.Text == "") //more efficient than creating an object then checking
                {
                    MessageBox.Show("Please make sure to fill out the Class and Assignment fields", "Error");
                }
                else
                {
                    schedule.Add(new Assignment(dtpDueDate.Value.Date, cmbClass.Text, txtAssignment.Text));
                    if (!cmbClass.Items.Contains(cmbClass.Text)) cmbClass.Items.Add(cmbClass.Text);

                    PrintToList();
                    txtAssignment.Clear();
                    txtAssignment.Focus();
                }
            }
        }

        private void PrintToList() //sorts list by date and then prints it.
        {
            schedule.Sort((a, b) => 2 * DateTime.Compare(a.Date, b.Date) + a.ClassCode.CompareTo(b.ClassCode)); // less memory usage sorting in-place than creating another list to sort
            lstAssignmentsBox.Items.Clear();

            foreach (var item in schedule)
            {
                lstAssignmentsBox.Items.Add(item.Date.ToString("MM/dd/yyyy").PadRight(15) + item.ClassCode.PadRight(26) + item.AssignmentName.PadRight(55));
            }
        }

        private void MnuUploadText_Click(object sender, EventArgs e)
        {
            DialogResult textConfirmation = MessageBox.Show("To use this function the information must be stored in a similar format as the list (mm/dd/yyyy;class;assignnment;) in a text document. " +
                "\n\n Do you want to continue? ", "Upload from Text Files", MessageBoxButtons.YesNo);

            if (textConfirmation == DialogResult.Yes)
            {
                MessageBox.Show("Please select a text file to upload from", "Upload from Text Files");

                OpenFileDialog textFileOpen = new OpenFileDialog
                {
                    Title = "Upload from Text Files",
                    Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*",
                    InitialDirectory = @"C:\"
                };

                if (textFileOpen.ShowDialog() == DialogResult.OK)
                {
                    string textFileName = textFileOpen.FileName;
                    try
                    {
                        StreamReader inputText;
                        string textContents;

                        inputText = File.OpenText(textFileName);

                        while ((textContents = inputText.ReadLine()) != null)
                        {
                            string[] parts = textContents.Split(';');
                            int i = 0;
                            while (i < parts.Length - 1)
                            {
                                schedule.Add(new Assignment(DateTime.Parse(parts[i + 0].Trim()), parts[i + 1].Trim(), parts[i + 2].Trim()));
                                i += 3;
                            }
                        }

                        inputText.Close();
                        PrintToList();
                        MessageBox.Show("Assignments from text file uploaded successfully.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Invalid entry. Make sure data format is correct and that the selected file is valid.", "Error");
                        Console.WriteLine(ex.ToString());
                    }
                }
            }
        }

        private void MnuUploadExcel_Click(object sender, EventArgs e)
        {
            DialogResult spreadsheetConfirmation = MessageBox.Show("Data must be stored in a similar format as as a created schedule (e.g headers of Dates/Class/Assignment). " +
                "\n\n Do you want to continue? ", "Upload from Excel Files", MessageBoxButtons.YesNo);

            if (spreadsheetConfirmation == DialogResult.Yes)
            {
                MessageBox.Show("Please select an Excel file to upload from", "Upload from Excel files");

                OpenFileDialog excelFileOpen = new OpenFileDialog
                {
                    Title = "Upload from Excel file",
                    Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*",
                    InitialDirectory = @"C:\"
                };

                if (excelFileOpen.ShowDialog() == DialogResult.OK)
                {
                    string excelFileName = excelFileOpen.FileName;      // ***This value is saved for later overwriting/updating***
                    try
                    {
                        IXLWorkbook sourceWbook = new XLWorkbook(excelFileName);
                        var ws1 = sourceWbook.Worksheet("Assignments List");

                        int lastRow = ws1.LastRowUsed().RowNumber();
                        //string classID, assignment;
                        //DateTime dueDate;

                        for (int i = 0; i < lastRow - 1; i++)
                        {
                            //dueDate = DateTime.Parse(ws1.Cell(i + 2, 1).Value.ToString());
                            //classID = ws1.Cell(i + 2, 2).GetString();
                            //assignment = ws1.Cell(i + 2, 3).GetString();

                            schedule.Add(new Assignment(
                                DateTime.Parse(ws1.Cell(i + 2, 1).Value.ToString()),
                                ws1.Cell(i + 2, 2).GetString(),
                                ws1.Cell(i + 2, 3).GetString()));
                        }
                        PrintToList();
                        MessageBox.Show("Assignments from excel file uploaded successfully.", "Upload from Excel files");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Invalid entry. Make sure data format is correct and that the selected file is valid.", "Error");
                        Console.WriteLine(ex.ToString());
                    }
                }
            }
        }
    }
}