
namespace Schedule_Planner
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.lstAssignmentsBox = new System.Windows.Forms.ListBox();
            this.txtAssignment = new System.Windows.Forms.TextBox();
            this.dtpDueDate = new System.Windows.Forms.DateTimePicker();
            this.lblClass = new System.Windows.Forms.Label();
            this.lblAssignment = new System.Windows.Forms.Label();
            this.lblDueDate = new System.Windows.Forms.Label();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnRemove = new System.Windows.Forms.Button();
            this.btnBuild = new System.Windows.Forms.Button();
            this.btnHelp = new System.Windows.Forms.Button();
            this.cmbClass = new System.Windows.Forms.ComboBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.menFileImport = new System.Windows.Forms.MenuStrip();
            this.tsmFileImport = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuImportText = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuImportExcel = new System.Windows.Forms.ToolStripMenuItem();
            this.menFileImport.SuspendLayout();
            this.SuspendLayout();
            // 
            // lstAssignmentsBox
            // 
            this.lstAssignmentsBox.Font = new System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.lstAssignmentsBox.FormattingEnabled = true;
            this.lstAssignmentsBox.ItemHeight = 16;
            this.lstAssignmentsBox.Location = new System.Drawing.Point(26, 100);
            this.lstAssignmentsBox.Name = "lstAssignmentsBox";
            this.lstAssignmentsBox.ScrollAlwaysVisible = true;
            this.lstAssignmentsBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.lstAssignmentsBox.Size = new System.Drawing.Size(731, 244);
            this.lstAssignmentsBox.TabIndex = 4;
            this.lstAssignmentsBox.TabStop = false;
            // 
            // txtAssignment
            // 
            this.txtAssignment.Location = new System.Drawing.Point(357, 61);
            this.txtAssignment.Name = "txtAssignment";
            this.txtAssignment.Size = new System.Drawing.Size(400, 23);
            this.txtAssignment.TabIndex = 3;
            this.txtAssignment.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtAssignment_KeyPress);
            // 
            // dtpDueDate
            // 
            this.dtpDueDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpDueDate.Location = new System.Drawing.Point(26, 61);
            this.dtpDueDate.Name = "dtpDueDate";
            this.dtpDueDate.Size = new System.Drawing.Size(119, 23);
            this.dtpDueDate.TabIndex = 1;
            // 
            // lblClass
            // 
            this.lblClass.AutoSize = true;
            this.lblClass.Location = new System.Drawing.Point(151, 43);
            this.lblClass.Name = "lblClass";
            this.lblClass.Size = new System.Drawing.Size(34, 15);
            this.lblClass.TabIndex = 4;
            this.lblClass.Text = "Class";
            // 
            // lblAssignment
            // 
            this.lblAssignment.AutoSize = true;
            this.lblAssignment.Location = new System.Drawing.Point(357, 43);
            this.lblAssignment.Name = "lblAssignment";
            this.lblAssignment.Size = new System.Drawing.Size(70, 15);
            this.lblAssignment.TabIndex = 5;
            this.lblAssignment.Text = "Assignment";
            // 
            // lblDueDate
            // 
            this.lblDueDate.AutoSize = true;
            this.lblDueDate.Location = new System.Drawing.Point(26, 43);
            this.lblDueDate.Name = "lblDueDate";
            this.lblDueDate.Size = new System.Drawing.Size(55, 15);
            this.lblDueDate.TabIndex = 6;
            this.lblDueDate.Text = "Due Date";
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(26, 381);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(139, 46);
            this.btnAdd.TabIndex = 5;
            this.btnAdd.TabStop = false;
            this.btnAdd.Text = "Add";
            this.btnAdd.UseVisualStyleBackColor = true;
            this.btnAdd.Click += new System.EventHandler(this.BtnAdd_Click);
            // 
            // btnRemove
            // 
            this.btnRemove.Location = new System.Drawing.Point(219, 381);
            this.btnRemove.Name = "btnRemove";
            this.btnRemove.Size = new System.Drawing.Size(139, 46);
            this.btnRemove.TabIndex = 6;
            this.btnRemove.TabStop = false;
            this.btnRemove.Text = "Remove";
            this.btnRemove.UseVisualStyleBackColor = true;
            this.btnRemove.Click += new System.EventHandler(this.BtnRemove_Click);
            // 
            // btnBuild
            // 
            this.btnBuild.Location = new System.Drawing.Point(427, 381);
            this.btnBuild.Name = "btnBuild";
            this.btnBuild.Size = new System.Drawing.Size(139, 46);
            this.btnBuild.TabIndex = 7;
            this.btnBuild.TabStop = false;
            this.btnBuild.Text = "Build";
            this.btnBuild.UseVisualStyleBackColor = false;
            this.btnBuild.Click += new System.EventHandler(this.BtnBuild_Click);
            // 
            // btnHelp
            // 
            this.btnHelp.Location = new System.Drawing.Point(618, 381);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(139, 46);
            this.btnHelp.TabIndex = 8;
            this.btnHelp.TabStop = false;
            this.btnHelp.Text = "Help";
            this.btnHelp.UseVisualStyleBackColor = true;
            this.btnHelp.Click += new System.EventHandler(this.BtnHelp_Click);
            // 
            // cmbClass
            // 
            this.cmbClass.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cmbClass.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cmbClass.FormattingEnabled = true;
            this.cmbClass.Location = new System.Drawing.Point(151, 61);
            this.cmbClass.Name = "cmbClass";
            this.cmbClass.Size = new System.Drawing.Size(200, 23);
            this.cmbClass.Sorted = true;
            this.cmbClass.TabIndex = 2;
            // 
            // menFileImport
            // 
            this.menFileImport.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmFileImport});
            this.menFileImport.Location = new System.Drawing.Point(0, 0);
            this.menFileImport.Name = "menFileImport";
            this.menFileImport.Size = new System.Drawing.Size(800, 24);
            this.menFileImport.TabIndex = 10;
            this.menFileImport.Text = "menuStrip1";
            // 
            // tsmFileImport
            // 
            this.tsmFileImport.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuImportText,
            this.mnuImportExcel});
            this.tsmFileImport.Image = global::Schedule_Planner.Properties.Resources.import_file_symbol;
            this.tsmFileImport.Name = "tsmFileImport";
            this.tsmFileImport.Size = new System.Drawing.Size(92, 20);
            this.tsmFileImport.Text = "File Import";
            // 
            // mnuImportText
            // 
            this.mnuImportText.Image = global::Schedule_Planner.Properties.Resources.txt_file_symbol;
            this.mnuImportText.Name = "mnuImportText";
            this.mnuImportText.Size = new System.Drawing.Size(190, 22);
            this.mnuImportText.Text = "Import from Text File";
            this.mnuImportText.Click += new System.EventHandler(this.MnuImportText_Click);
            // 
            // mnuImportExcel
            // 
            this.mnuImportExcel.Image = global::Schedule_Planner.Properties.Resources.xl_file_symbol;
            this.mnuImportExcel.Name = "mnuImportExcel";
            this.mnuImportExcel.Size = new System.Drawing.Size(190, 22);
            this.mnuImportExcel.Text = "Import from Excel File";
            this.mnuImportExcel.Click += new System.EventHandler(this.MnuImportExcel_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.cmbClass);
            this.Controls.Add(this.btnHelp);
            this.Controls.Add(this.btnBuild);
            this.Controls.Add(this.btnRemove);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.lblDueDate);
            this.Controls.Add(this.lblAssignment);
            this.Controls.Add(this.lblClass);
            this.Controls.Add(this.dtpDueDate);
            this.Controls.Add(this.txtAssignment);
            this.Controls.Add(this.lstAssignmentsBox);
            this.Controls.Add(this.menFileImport);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menFileImport;
            this.Name = "Form1";
            this.Text = "Schedule Planner";
            this.menFileImport.ResumeLayout(false);
            this.menFileImport.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox lstAssignmentsBox;
        private System.Windows.Forms.TextBox txtAssignment;
        private System.Windows.Forms.DateTimePicker dtpDueDate;
        private System.Windows.Forms.Label lblClass;
        private System.Windows.Forms.Label lblAssignment;
        private System.Windows.Forms.Label lblDueDate;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnRemove;
        private System.Windows.Forms.Button btnBuild;
        private System.Windows.Forms.Button btnHelp;
        private System.Windows.Forms.ComboBox cmbClass;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.MenuStrip menFileImport;
        private System.Windows.Forms.ToolStripMenuItem tsmFileImport;
        private System.Windows.Forms.ToolStripMenuItem mnuImportText;
        private System.Windows.Forms.ToolStripMenuItem mnuImportExcel;
    }
}

