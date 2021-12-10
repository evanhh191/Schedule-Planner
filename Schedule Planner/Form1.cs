using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace Schedule_Planner
{
    public partial class Form1 : Form
    {
        private readonly List<Assignment> schedule = new List<Assignment>();      //hidden list of Assignment class, used to store info and to make strings for displayed list

        public Form1()
        {
            InitializeComponent();
        }

        private void BtnAdd_Click(object sender, EventArgs e)    //adds the info from the date, class, and time fields to the list and resorts it by date then classCode
        {
            if (cmbClass.Text == "" || txtAssignment.Text == "")
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

        private void BtnRemove_Click(object sender, EventArgs e) //Removes selected item(s) from list.
        {
            if (lstAssignmentsBox.SelectedItems.Count == 0) MessageBox.Show("Please select an item on the list to remove.\n", "Error");
            else
            {
                var indices = new List<int>();
                foreach (var item in lstAssignmentsBox.SelectedItems)
                {
                    indices.Add(lstAssignmentsBox.Items.IndexOf(item));
                }
                for (int i = indices.Count - 1; i >= 0; i--)
                {
                    schedule.RemoveAt(indices[i]);
                }
                PrintToList();
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
            if (schedule.Count == 0) MessageBox.Show("Please ensure that there is at least one entry in the list", "Error");
            else
            {
                int dateRange = (int)(schedule[^1].Date.ToOADate() - schedule[0].Date.ToOADate() + 1); //converting to serialized date. Otherwise doesn't work. adding one to be inclusive of start date

                DialogResult buildResult = MessageBox.Show("Are you ready to create an Excel calendar with the given data?" +
                    "\nYour calendar will range from: " + schedule[0].Date.ToShortDateString() + " to: " + schedule[^1].Date.ToShortDateString()
                         + "\nFor a date range of: " + dateRange + " day(s)" + "\nNumber of assignments: " + schedule.Count, "Build", MessageBoxButtons.YesNo);

                if (buildResult == DialogResult.No) { }
                else if (buildResult == DialogResult.Yes)       // ** THIS IS WHERE THE SPREADSHEET BUILDING WILL HAPPEN **
                {
                    FolderBrowserDialog browserDialog = new FolderBrowserDialog();
                    string fileName;

                    if (browserDialog.ShowDialog() == DialogResult.OK)
                    {
                        fileName = browserDialog.SelectedPath;

                        DialogResult locationResult = MessageBox.Show("Please close spreadsheet if open. \n\n\nSave to: " + fileName + " ?", "Build", MessageBoxButtons.OKCancel);

                        if (locationResult == DialogResult.OK)
                        {
                            BuildSpreadsheet(fileName, dateRange);
                        }
                    }
                }
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
                        .Fill.SetBackgroundColor(XLColor.FromArgb(255, 230, 153))
                        .Border.SetOutsideBorder(XLBorderStyleValues.Thick);
                    worksheet2.Cell(1, 2 * i + 1).Value = (DayOfWeek)i;
                }

                int formulaRange = schedule.Count;
                Boolean first = true;

                for (int i = (int)schedule[0].Date.DayOfWeek; i < dateRange + (int)schedule[0].Date.DayOfWeek; i++)
                {
                    //(i/7) * 7 counts the number of weeks so far. ex day 6, which is a saturday, is week 0 because for ints, 6/7 = 0
                    int rowIncrementer = (i / 7) * 7;
                    //green date header
                    worksheet2.Range(worksheet2.Cell(2 + rowIncrementer, 2 * (i % 7) + 1), worksheet2.Cell(2 + rowIncrementer, 2 * (i % 7) + 2))
                        .Merge().Style
                        .Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                        .Fill.SetBackgroundColor(XLColor.LightGreen)
                        .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                        .NumberFormat.SetFormat("m/d");
                    //blue background for assignments and class codes
                    worksheet2.Range(worksheet2.Cell(3 + rowIncrementer, 2 * (i % 7) + 1), worksheet2.Cell(3 + rowIncrementer + 5, 2 * (i % 7) + 2))
                        .Style.Fill.SetBackgroundColor(XLColor.FromArgb(221, 235, 247));
                    //thick border outside entire day block
                    worksheet2.Range(worksheet2.Cell(2 + rowIncrementer, 2 * (i % 7) + 1), worksheet2.Cell(3 + rowIncrementer + 5, 2 * (i % 7) + 2))
                        .Style.Border.SetOutsideBorder(XLBorderStyleValues.Thick);
                    //thin border to separate class codes and assignment names
                    worksheet2.Range(worksheet2.Cell(3 + rowIncrementer, 2 * (i % 7) + 1), worksheet2.Cell(3 + rowIncrementer + 5, 2 * (i % 7) + 1))
                        .Style.Border.SetRightBorder(XLBorderStyleValues.Thin);
                    var currentCell = worksheet2.Cell(2 + rowIncrementer, 2 * (i % 7) + 1);
                    //first date that other dates are based off of
                    if (first)
                    {
                        currentCell.Value = schedule[i - (int)schedule[0].Date.DayOfWeek].Date.ToShortDateString();
                        first = false;
                    }
                    //dates for first and second weeks, aside from first days in week
                    else if (i / 7 < 1 || i / 7 == 1 && i % 7 != 0)
                    {
                        currentCell.FormulaR1C1 = "=RC[-2]+1";
                    }
                    //date for first day of second week. needed if first day of first week starts in middle of week
                    else if (i / 7 == 1 && i % 7 == 0)
                    {
                        currentCell.FormulaR1C1 = "=R[-7]C[12] + 1";
                    }
                    //dates for every day after the first two weeks
                    else
                    {
                        currentCell.FormulaR1C1 = "=R[-7]C+7";
                    }
                    for (int j = 0; j < 6; j++)
                    {
                        //class code formula
                        worksheet2.Cell(3 + rowIncrementer + j, 2 * (i % 7) + 1).FormulaR1C1 = "IF(COUNTIF(Sheet1!R2C1:R" + (formulaRange + 1) + "C1,R[-" + (j + 1) + "]C)>" + j
                            + ",INDEX(Sheet1!R2C1:Sheet1!R" + (formulaRange + 1) + "C3,MATCH(R[-" + (j + 1) + "]C,Sheet1!R2C1:R" + (formulaRange + 1) + "C1,0)+" + j + ",2),\"\")";
                        //assignment name formula
                        worksheet2.Cell(3 + rowIncrementer + j, 2 * (i % 7) + 2).FormulaR1C1 = "IF(COUNTIF(Sheet1!R2C1:R" + (formulaRange + 1) + "C1,R[-" + (j + 1) + "]C[-1])>" + j
                            + ",INDEX(Sheet1!R2C1:Sheet1!R" + (formulaRange + 1) + "C3,MATCH(R[-" + (j + 1) + "]C[-1],Sheet1!R2C1:R" + (formulaRange + 1) + "C1,0)+" + j + ",3),\"\")";
                    }
                }
                worksheet1.Rows().AdjustToContents();
                worksheet1.Columns().AdjustToContents();
                worksheet2.Rows().AdjustToContents();
                worksheet2.Columns().AdjustToContents();

                worksheet1.SheetView.FreezeRows(1);
                worksheet2.SheetView.FreezeRows(1);
                workbook.SaveAs(fileName);
                MessageBox.Show("A spreadsheet calendar has been created at: " + fileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Build failed. Check if file is open and try again.\nDetails: " + ex.ToString(), "Error");
            }
        }

        /* Saving for later developement/quick enable fill button
        private void Button3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 160; i++)
            {
                schedule.Add(new Assignment(DateTime.Now.AddDays(i % 40), "Test" + i % 4, "test assignment name" + i));
            }
            for (int i = 160; i < 234; i++)
            {
                schedule.Add(new Assignment(DateTime.Now.AddDays(i), "Test" + i % 4, "test assignment name" + i));
            }
            PrintToList();
        }
        */

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
            DialogResult textConfirmation = MessageBox.Show("To use this function, the information must be stored in a similar format as the list (mm/dd/yyyy;class;assignnment;) in a text document. " +
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
                    try
                    {
                        StreamReader inputText;
                        string textContents;

                        inputText = File.OpenText(textFileOpen.FileName);

                        while ((textContents = inputText.ReadLine()) != null)
                        {
                            string[] parts = textContents.Split(';');
                            int i = 0;
                            while (i < parts.Length - 1)
                            {
                                schedule.Add(new Assignment(DateTime.Parse(parts[i + 0].Trim()), parts[i + 1].Trim(), parts[i + 2].Trim()));
                                if (!cmbClass.Items.Contains(parts[i + 1])) cmbClass.Items.Add(parts[i + 1]);
                                i += 3;
                            }
                        }

                        inputText.Close();
                        PrintToList();
                        MessageBox.Show("Assignments from text file uploaded successfully.");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Invalid entry. Make sure data format is correct and that the selected file is valid.\nDetails: " + ex.ToString(), "Error");
                    }
                }
            }
        }

        private void MnuUploadExcel_Click(object sender, EventArgs e)
        {
            DialogResult spreadsheetConfirmation = MessageBox.Show("Data must be stored in a similar format as as a created schedule (e.g headers of Dates/Class/Assignment). "
                + "\nThe leftmost sheet must contain the list of dates."
                + "\n\n Do you want to continue? ", "Upload from Excel Files", MessageBoxButtons.YesNo);

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
                    try
                    {
                        IXLWorkbook sourceWbook = new XLWorkbook(excelFileOpen.FileName);
                        var ws1 = sourceWbook.Worksheet(1);

                        int lastRow = ws1.LastRowUsed().RowNumber();

                        for (int i = 0; i < lastRow - 1; i++)
                        {
                            schedule.Add(new Assignment(
                                DateTime.Parse(ws1.Cell(i + 2, 1).Value.ToString()),
                                ws1.Cell(i + 2, 2).GetString(),
                                ws1.Cell(i + 2, 3).GetString()));

                            if (!cmbClass.Items.Contains(ws1.Cell(i + 2, 2).GetString())) cmbClass.Items.Add(ws1.Cell(i + 2, 2).GetString());
                        }
                        PrintToList();
                        sourceWbook.Dispose();
                        MessageBox.Show("Assignments from excel file uploaded successfully.", "Upload from Excel files");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Invalid entry. Make sure data format is correct and that the selected file is valid.\nDetails: " + ex.ToString(), "Error");
                    }
                }
            }
        }
    }
}