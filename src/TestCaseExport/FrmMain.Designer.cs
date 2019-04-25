namespace TestCaseExport
{
    partial class FrmMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMain));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.comBoxTestPlan = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.treeView_suite = new System.Windows.Forms.TreeView();
            this.lblTestPlan = new System.Windows.Forms.Label();
            this.btnTeamProject = new System.Windows.Forms.Button();
            this.txtTeamProject = new System.Windows.Forms.TextBox();
            this.lblTeamProject = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnFolderBrowse = new System.Windows.Forms.Button();
            this.txtSaveFolder = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.btnHelp = new System.Windows.Forms.Button();
            this.SeparateSheets = new System.Windows.Forms.CheckBox();
            this.ExportResults = new System.Windows.Forms.CheckBox();
            this.NoSubSuite = new System.Windows.Forms.CheckBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.comBoxTestPlan);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.treeView_suite);
            this.groupBox1.Controls.Add(this.lblTestPlan);
            this.groupBox1.Controls.Add(this.btnTeamProject);
            this.groupBox1.Controls.Add(this.txtTeamProject);
            this.groupBox1.Controls.Add(this.lblTeamProject);
            this.groupBox1.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(10, 75);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(244, 369);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Source";
            // 
            // comBoxTestPlan
            // 
            this.comBoxTestPlan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comBoxTestPlan.DropDownWidth = 200;
            this.comBoxTestPlan.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comBoxTestPlan.FormattingEnabled = true;
            this.comBoxTestPlan.Location = new System.Drawing.Point(11, 98);
            this.comBoxTestPlan.Name = "comBoxTestPlan";
            this.comBoxTestPlan.Size = new System.Drawing.Size(215, 23);
            this.comBoxTestPlan.TabIndex = 2;
            this.comBoxTestPlan.SelectedIndexChanged += new System.EventHandler(this.comBoxTestPlan_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(8, 139);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(61, 15);
            this.label4.TabIndex = 12;
            this.label4.Text = "Test Suite:";
            // 
            // treeView_suite
            // 
            this.treeView_suite.AllowDrop = true;
            this.treeView_suite.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.treeView_suite.HideSelection = false;
            this.treeView_suite.Location = new System.Drawing.Point(11, 160);
            this.treeView_suite.Name = "treeView_suite";
            this.treeView_suite.Size = new System.Drawing.Size(217, 200);
            this.treeView_suite.TabIndex = 9;
            // 
            // lblTestPlan
            // 
            this.lblTestPlan.AutoSize = true;
            this.lblTestPlan.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTestPlan.Location = new System.Drawing.Point(8, 75);
            this.lblTestPlan.Name = "lblTestPlan";
            this.lblTestPlan.Size = new System.Drawing.Size(59, 15);
            this.lblTestPlan.TabIndex = 0;
            this.lblTestPlan.Text = "Test Plan:";
            this.lblTestPlan.Click += new System.EventHandler(this.lblTestPlan_Click);
            // 
            // btnTeamProject
            // 
            this.btnTeamProject.Location = new System.Drawing.Point(191, 43);
            this.btnTeamProject.Name = "btnTeamProject";
            this.btnTeamProject.Size = new System.Drawing.Size(35, 23);
            this.btnTeamProject.TabIndex = 1;
            this.btnTeamProject.Text = "...";
            this.btnTeamProject.UseVisualStyleBackColor = true;
            this.btnTeamProject.Click += new System.EventHandler(this.btnTeamProject_Click);
            // 
            // txtTeamProject
            // 
            this.txtTeamProject.BackColor = System.Drawing.Color.White;
            this.txtTeamProject.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtTeamProject.ForeColor = System.Drawing.SystemColors.GrayText;
            this.txtTeamProject.Location = new System.Drawing.Point(11, 42);
            this.txtTeamProject.Name = "txtTeamProject";
            this.txtTeamProject.ReadOnly = true;
            this.txtTeamProject.Size = new System.Drawing.Size(172, 24);
            this.txtTeamProject.TabIndex = 0;
            this.txtTeamProject.TextChanged += new System.EventHandler(this.txtTeamProject_TextChanged);
            // 
            // lblTeamProject
            // 
            this.lblTeamProject.AutoSize = true;
            this.lblTeamProject.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTeamProject.Location = new System.Drawing.Point(8, 21);
            this.lblTeamProject.Name = "lblTeamProject";
            this.lblTeamProject.Size = new System.Drawing.Size(89, 15);
            this.lblTeamProject.TabIndex = 0;
            this.lblTeamProject.Text = "Connect to TFS:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.txtFileName);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.btnFolderBrowse);
            this.groupBox2.Controls.Add(this.txtSaveFolder);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(260, 75);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(378, 144);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Destination";
            // 
            // txtFileName
            // 
            this.txtFileName.BackColor = System.Drawing.Color.White;
            this.txtFileName.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFileName.Location = new System.Drawing.Point(11, 98);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(323, 24);
            this.txtFileName.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(8, 75);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(109, 15);
            this.label2.TabIndex = 0;
            this.label2.Text = "Name of Excel File:";
            // 
            // btnFolderBrowse
            // 
            this.btnFolderBrowse.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFolderBrowse.Location = new System.Drawing.Point(253, 38);
            this.btnFolderBrowse.Name = "btnFolderBrowse";
            this.btnFolderBrowse.Size = new System.Drawing.Size(81, 28);
            this.btnFolderBrowse.TabIndex = 4;
            this.btnFolderBrowse.Text = "Browse...";
            this.btnFolderBrowse.UseVisualStyleBackColor = true;
            this.btnFolderBrowse.Click += new System.EventHandler(this.btnFolderBrowse_Click);
            // 
            // txtSaveFolder
            // 
            this.txtSaveFolder.BackColor = System.Drawing.Color.White;
            this.txtSaveFolder.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSaveFolder.ForeColor = System.Drawing.SystemColors.GrayText;
            this.txtSaveFolder.Location = new System.Drawing.Point(11, 42);
            this.txtSaveFolder.Name = "txtSaveFolder";
            this.txtSaveFolder.ReadOnly = true;
            this.txtSaveFolder.Size = new System.Drawing.Size(236, 24);
            this.txtSaveFolder.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(8, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(172, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "Location to Save Exported File:";
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(556, 410);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(82, 25);
            this.btnCancel.TabIndex = 7;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnExport
            // 
            this.btnExport.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExport.Location = new System.Drawing.Point(453, 410);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(82, 25);
            this.btnExport.TabIndex = 6;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.richTextBox1);
            this.groupBox3.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.Location = new System.Drawing.Point(10, 3);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(628, 66);
            this.groupBox3.TabIndex = 0;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Welcome";
            // 
            // richTextBox1
            // 
            this.richTextBox1.BackColor = System.Drawing.SystemColors.HighlightText;
            this.richTextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.richTextBox1.Enabled = false;
            this.richTextBox1.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.richTextBox1.ForeColor = System.Drawing.SystemColors.MenuText;
            this.richTextBox1.Location = new System.Drawing.Point(6, 24);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None;
            this.richTextBox1.Size = new System.Drawing.Size(607, 36);
            this.richTextBox1.TabIndex = 15;
            this.richTextBox1.Text = "This tool can be used to export test cases from TFS to Microsoft excel. \nIf you a" +
    "re exporting a large number of test cases, the export process may take some time" +
    ".\n";
            this.richTextBox1.TextChanged += new System.EventHandler(this.richTextBox1_TextChanged);
            // 
            // btnHelp
            // 
            this.btnHelp.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnHelp.Location = new System.Drawing.Point(260, 410);
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.Size = new System.Drawing.Size(82, 25);
            this.btnHelp.TabIndex = 8;
            this.btnHelp.Text = "About";
            this.btnHelp.UseVisualStyleBackColor = true;
            this.btnHelp.Click += new System.EventHandler(this.btnAbout_Click);
            // 
            // SeparateSheets
            // 
            this.SeparateSheets.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SeparateSheets.Location = new System.Drawing.Point(11, 89);
            this.SeparateSheets.Name = "SeparateSheets";
            this.SeparateSheets.Size = new System.Drawing.Size(323, 32);
            this.SeparateSheets.TabIndex = 0;
            this.SeparateSheets.Text = "Export Each Test Suite into Separate Sheets";
            this.SeparateSheets.UseVisualStyleBackColor = true;
            this.SeparateSheets.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // ExportResults
            // 
            this.ExportResults.AutoSize = true;
            this.ExportResults.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ExportResults.Location = new System.Drawing.Point(11, 137);
            this.ExportResults.Name = "ExportResults";
            this.ExportResults.Size = new System.Drawing.Size(128, 19);
            this.ExportResults.TabIndex = 13;
            this.ExportResults.Text = "Export Test Results";
            this.ExportResults.UseVisualStyleBackColor = true;
            // 
            // NoSubSuite
            // 
            this.NoSubSuite.AutoSize = true;
            this.NoSubSuite.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.NoSubSuite.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.NoSubSuite.Location = new System.Drawing.Point(11, 25);
            this.NoSubSuite.Name = "NoSubSuite";
            this.NoSubSuite.Size = new System.Drawing.Size(309, 49);
            this.NoSubSuite.TabIndex = 14;
            this.NoSubSuite.Text = "\r\nExport the Selected Suite Only.\r\n(Test cases from the sub suites will not be ex" +
    "ported)";
            this.NoSubSuite.UseVisualStyleBackColor = true;
            this.NoSubSuite.CheckedChanged += new System.EventHandler(this.checkBox2_CheckedChanged);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.NoSubSuite);
            this.groupBox4.Controls.Add(this.ExportResults);
            this.groupBox4.Controls.Add(this.SeparateSheets);
            this.groupBox4.Font = new System.Drawing.Font("Calibri", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox4.Location = new System.Drawing.Point(260, 225);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(378, 174);
            this.groupBox4.TabIndex = 8;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Export Options";
            // 
            // FrmMain
            // 
            this.AcceptButton = this.btnExport;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(650, 447);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.btnHelp);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Calibri", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "FrmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Export Test Cases from TFS to Excel";
            this.Load += new System.EventHandler(this.FrmMain_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lblTestPlan;
        private System.Windows.Forms.Button btnTeamProject;
        private System.Windows.Forms.TextBox txtTeamProject;
        private System.Windows.Forms.Label lblTeamProject;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnFolderBrowse;
        private System.Windows.Forms.TextBox txtSaveFolder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btnHelp;
        private System.Windows.Forms.TreeView treeView_suite;
        private System.Windows.Forms.ComboBox comBoxTestPlan;
        private System.Windows.Forms.CheckBox SeparateSheets;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox ExportResults;
        private System.Windows.Forms.CheckBox NoSubSuite;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.GroupBox groupBox4;
    }
}

