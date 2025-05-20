using System;
using System.Drawing;
using System.Windows.Forms;

using System.Text.RegularExpressions;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Runtime.Remoting.Messaging;
using Microsoft.Office.Interop.Excel;
using Microsoft.TeamFoundation.TestManagement.WebApi;

namespace TFSTestRunExport
{
    public partial class FrmMain : Form
    {
        private TfsTeamProjectCollection _tfs;
        private ITestManagementTeamProject _teamProject;
        private ITestPlanCollection plans;
        private ITestPlan plan;
        private ITestCaseCollection testCases;
        //private ITestSuiteEntry suite;
        private WorkItemStore _store = null;
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Excel.Worksheet xlWSheet;
        object misValue = System.Reflection.Missing.Value;
        Excel.Range chartRange;
        int row = 2;
        int upperBound = 0;
        int lowerBound = 0;
        int sheetno = 1;
        int defaultSheets;
        int sheetcount;
        string Sname;
        String unfixedBugs = "";
        private Dictionary<string, int> testMetrics = new Dictionary<string, int>() {
            { "Total", 0 },
            { "Executed", 0 },
            { "Passed", 0 },
            { "Failed", 0 }
        };
        
        private Dictionary<string, HashSet<int>> bugMetrics = new Dictionary<string, HashSet<int>>() {
            { "Fixed", new HashSet<int>() },
            { "Deferred", new HashSet<int>() }
        };

        private void resetMetrics() {
            testMetrics["Total"] = 0;
            testMetrics["Executed"] = 0;
            testMetrics["Passed"] = 0;
            testMetrics["Failed"] = 0;

            bugMetrics["Fixed"] = new HashSet<int>();
            bugMetrics["Deferred"] = new HashSet<int>();
        }

        private delegate void Execute();

        public FrmMain()
        {
            InitializeComponent();
        }

        private void btnTeamProject_Click(object sender, EventArgs e)
        {
            //Displaying the Team Project selection dialog to select the desired team project.
            TeamProjectPicker tpp = new TeamProjectPicker(TeamProjectPickerMode.SingleProject, false);
            tpp.ShowDialog();

            //Following actions will be executed only if a team project is selected in the the opened dialog.
            if (tpp.SelectedTeamProjectCollection != null)
            {
                this._tfs = tpp.SelectedTeamProjectCollection;
                ITestManagementService test_service = (ITestManagementService)_tfs.GetService(typeof(ITestManagementService));
                _store = (WorkItemStore)_tfs.GetService(typeof(WorkItemStore));
                this._teamProject = test_service.GetTeamProject(tpp.SelectedProjects[0].Name);

                //Populating the text field Team Project name (txtTeamProject) with the name of the selected team project.
                txtTeamProject.Text = tpp.SelectedProjects[0].Name;

                //Call to method "Get_TestPlans" to get the test plans in the selected team project
                Get_TestPlans(_teamProject);
            }

        }

        private void Get_TestPlans(ITestManagementTeamProject teamProject)
        {
            //Getting all the test plans in the collection "plans" using the query.
            this.plans = teamProject.TestPlans.Query("Select * From TestPlan");
            comBoxTestPlan.Items.Clear();

            treeView_suite.BackColor = Color.White;

            foreach (ITestPlan plan in plans)
            {
                //Populating the plan selection dropdown list with the name of Test Plans in the selected team project.
                comBoxTestPlan.Items.Add(plan.Name);
            }

        }

        //Following method is invoked whenever a Test Plan is selected in the dropdown list.
        //Acording to the selected Test Plan in the dropdown list the Test Suites present in the selected Test Plan are populated in the Test Suite selection dropdown.
        private void comBoxTestPlan_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            treeView_suite.Nodes.Clear();
            int i = -1;
            if (comBoxTestPlan.SelectedIndex >= 0)
            {
                i = comBoxTestPlan.SelectedIndex;
                this.plan = plans[i];
                this.Cursor = Cursors.Arrow;
                if (plans[i].RootSuite != null)
                {
                    TreeNode rootnode = new TreeNode();
                    rootnode.Name = plans[i].RootSuite.Id.ToString();
                    rootnode.Text = plans[i].RootSuite.Title.ToString();
                    treeView_suite.Nodes.Add(rootnode);
                    if (plans[i].RootSuite.SubSuites != null && plans[i].RootSuite.SubSuites.Count > 0)
                    {
                        Get_subsuites(plans[i].RootSuite, rootnode);
                    }


                }

            }

        }

        private void Get_subsuites(IStaticTestSuite rootsuite1, TreeNode node1)
        {
            ITestSuiteCollection subsuite1 = rootsuite1.SubSuites;

            foreach (ITestSuiteBase suite in subsuite1)
            {

                if (suite != null)
                {
                    TreeNode subnode = new TreeNode();
                    subnode.Text = suite.Title.ToString();
                    subnode.Name = suite.Id.ToString();
                    node1.Nodes.Add(subnode);
                    if (suite.TestSuiteType == TestSuiteType.StaticTestSuite)
                    {
                        IStaticTestSuite subsuite2 = suite as IStaticTestSuite;
                        if (subsuite2 != null && (subsuite2.SubSuites.Count > 0))
                        {
                            Get_subsuites(subsuite2, subnode);
                        }
                    }

                }
            }

        }

        private void btnFolderBrowse_Click(object sender, EventArgs e)
        {
            folderBrowserDialog.ShowDialog();
            txtSaveFolder.Text = folderBrowserDialog.SelectedPath;
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (NoSubSuite.Checked == true)
            {
                SeparateSheets.Checked = false;
                SeparateSheets.Enabled = false;
            }
            if (NoSubSuite.Checked == false)
            {
                SeparateSheets.Enabled = true;
            }
        }

        private string removehtmltags(string text)
        {
            text = text.Replace("</P><P>", System.Environment.NewLine);


            text = Regex.Replace(text, "<.*?>", "");
            text = text.Replace("&#160;", "");

            return text;
        }

        private string formatsheetname(string str)
        {
            str = str.Replace("/", "");
            str = str.Replace("\\", "");
            str = str.Replace(":", "");
            str = str.Replace("?", "");
            str = str.Replace("[", "");
            str = str.Replace("]", "");
            str = str.Replace("*", "");

            if (str.Length > 30)
                str = str.Substring(0, 30);

            return str;
        }

        private void Get_TestCases(ITestSuiteBase testSuite)
        {
            this.testCases = testSuite.AllTestCases;
            if (NoSubSuite.Checked == true)
            {
                this.testCases.Clear();
                foreach (ITestSuiteEntry tse in testSuite.TestCases)
                {
                    if (tse.EntryType == TestSuiteEntryType.TestCase)
                    {
                        if (tse.TestCase != null)
                        {
                            testCases.Add(tse.TestCase);
                        }
                    }
                }
            }
        }

        private void midCenterCellRange(string x, string y)
        {
            chartRange = xlWorkSheet.get_Range(x, y);
            chartRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            chartRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
        }

        private void leftCenterCellRange(string x, string y)
        {
            chartRange = xlWorkSheet.get_Range(x, y);
            chartRange.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            chartRange.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
        }

        private void setResultBGColor(object cellObj, string result)
        {
            if(cellObj is Excel.Range cell)
            {
                switch(result)
                {
                    case "Passed":
                        cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(215, 230, 180));
                        break;
                    case "Failed":
                        cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(230,180,180));
                        break;
                    default:
                        return;
                }
            }
        }

        private void export(Excel.Worksheet xlWorkSheet, ITestCaseCollection testcases)
        {
            // Set headers on first row in sheet
            xlWorkSheet.Rows[1].RowHeight = 45;

            xlWorkSheet.Cells[1, 1] = "Test Case ID";
            xlWorkSheet.Cells[1, 2] = "Test Case Name";
            xlWorkSheet.Cells[1, 3] = "Final State:\n passed / failed";
            xlWorkSheet.Cells[1, 4] = "Run Number";
            xlWorkSheet.Cells[1, 5] = "Run Date";
            xlWorkSheet.Cells[1, 6] = "Configuration";
            xlWorkSheet.Cells[1, 7] = "Run Result";
            xlWorkSheet.Cells[1, 8] = "Bug IDs";
            xlWorkSheet.Cells[1, 9] = "Ran By";

            // Set column widths
            (xlWorkSheet.Columns["A", Type.Missing]).ColumnWidth = 11;
            (xlWorkSheet.Columns["B", Type.Missing]).ColumnWidth = 100;
            (xlWorkSheet.Columns["C", Type.Missing]).ColumnWidth = 15;
            (xlWorkSheet.Columns["D", Type.Missing]).ColumnWidth = 9;
            (xlWorkSheet.Columns["E", Type.Missing]).ColumnWidth = 15;
            (xlWorkSheet.Columns["F", Type.Missing]).ColumnWidth = 15;
            (xlWorkSheet.Columns["G", Type.Missing]).ColumnWidth = 15;
            (xlWorkSheet.Columns["H", Type.Missing]).ColumnWidth = 30;
            (xlWorkSheet.Columns["I", Type.Missing]).ColumnWidth = 30;

            Stopwatch querySW = new Stopwatch();
            querySW.Start();

            WorkItemStore workItemStore = (WorkItemStore)_tfs.GetService(typeof(WorkItemStore));

            var allResults = _teamProject.TestResults
                .Query(
                    "SELECT * FROM TestResult " +
                    "WHERE TestResult.TestPlanId = " + plan.Id
                )
                .ToLookup(r => r.TestCaseId);

            querySW.Stop();
            Console.WriteLine("Query Time Elapsed={0}", querySW.Elapsed);

            foreach (ITestCase testCase in testCases)
            {
                int i = 1;
                String finalOutcome = "";
                upperBound = row;

                testMetrics["Total"]++;

                #region ExportResults

                var testResultHistory = allResults[testCase.Id];

                Stopwatch bugSW = new Stopwatch();
                bugSW.Start();
                if (testResultHistory.Count() == 0)
                {
                    finalOutcome = "Not Run";
                    row++;
                }
                else
                {
                    foreach (ITestCaseResult result in testResultHistory)
                    {
                        bugSW.Stop();
                        Console.WriteLine("Get Next Result Time: " + bugSW.Elapsed);

                        String rez = removehtmltags(result.Outcome.ToString());
                        xlWorkSheet.Cells[row, 4] = i;
                        xlWorkSheet.Cells[row, 5] = removehtmltags(result.DateCompleted.Date.ToShortDateString());
                        xlWorkSheet.Cells[row, 6] = removehtmltags(result.TestConfigurationName.ToString());
                        xlWorkSheet.Cells[row, 7] = rez;
                        Excel.Range cell = xlWorkSheet.Cells[row, 7];
                        setResultBGColor(cell, rez);

                        var runBy = result.RunByName;
                        if (runBy != null)
                        {
                            runBy.ToString();
                            xlWorkSheet.Cells[row, 9] = removehtmltags(runBy);
                        }

                        if (rez == "Failed")
                        {
                            String bugs = "";
                            var associatedLinks = result.QueryAssociatedWorkItems();

                            foreach (int linkIter in associatedLinks)
                            {
                                var workItemBug = workItemStore.GetWorkItem(linkIter);
                                if (workItemBug.Type.Name.ToString() == "Bug")
                                {
                                    string bug = "#" + linkIter;
                                    bugs += bug + ", ";

                                    if (workItemBug.State == "Done") {
                                        bugMetrics["Fixed"].Add(workItemBug.Id);
                                    } else if (workItemBug.State != "Removed" 
                                            && !unfixedBugs.Contains(bug)) {
                                        unfixedBugs += (unfixedBugs.Equals("")) ? bug : ", " + bug;
                                        bugMetrics["Deferred"].Add(workItemBug.Id);
                                    }
                                }
                            }
                            if (!bugs.Equals(""))
                            {
                                bugs = Regex.Replace(bugs, ", $", "");
                                xlWorkSheet.Cells[row, 8] = bugs;
                            }
                        }
                        row++;
                        i++;
                        bugSW.Start();
                    }

                    finalOutcome = testResultHistory
                        .ElementAt(testResultHistory.Count() - 1)
                        .Outcome.ToString();
                    switch (finalOutcome) {
                        case "Passed":
                            testMetrics["Executed"]++;
                            testMetrics["Passed"]++;
                            break;
                        case "Failed":
                            testMetrics["Executed"]++;
                            testMetrics["Failed"]++;
                            break;
                    }

                }
                bugSW.Stop();
                #endregion

                lowerBound = (row - 1);
                chartRange = xlWorkSheet.get_Range("a" + upperBound, "i" + lowerBound);

                chartRange.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                chartRange.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
                chartRange.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
                chartRange.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
                chartRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

                chartRange = xlWorkSheet.get_Range("a" + upperBound, "a" + lowerBound);
                chartRange.Merge(false);
                chartRange.FormulaR1C1 = removehtmltags(testCase.Id.ToString());

                chartRange = xlWorkSheet.get_Range("b" + upperBound, "b" + lowerBound);
                chartRange.Merge(false);
                chartRange.FormulaR1C1 = testCase.Title.ToString();

                chartRange = xlWorkSheet.get_Range("c" + upperBound, "c" + lowerBound);
                chartRange.Merge(false);
                chartRange.FormulaR1C1 = finalOutcome;
                setResultBGColor(chartRange, finalOutcome);

                midCenterCellRange("a" + upperBound, "g" + lowerBound);
                leftCenterCellRange("h" + upperBound, "i" + lowerBound);
                leftCenterCellRange("b" + upperBound, "b" + lowerBound);
            }

            // Turn on Wrap text on entire table
            lowerBound = (row - 1);
            chartRange = xlWorkSheet.get_Range("a1", "I" + lowerBound);
            chartRange.Cells.WrapText = true;
            chartRange.Cells.Font.Name = "Calibri";

            // Draw border and align + bold header
            chartRange = xlWorkSheet.get_Range("a1", "I1");
            chartRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            chartRange.Font.Bold = true;
            midCenterCellRange("A1", "I1");

            row = 2;
            upperBound = 0;
            lowerBound = 0;
        }

        private void exportmultisheet(ITestSuiteBase itsb, Excel.Workbook xlBook)
        {
            // count bugs that are represented by Requirement Test Suites
            if (itsb.TestSuiteType == TestSuiteType.RequirementTestSuite) 
            {
                var reqSuite = itsb as IRequirementTestSuite;
                if (reqSuite != null && reqSuite.RequirementId > 0)
                {
                    WorkItem linkedItem = _store.GetWorkItem(reqSuite.RequirementId);
                    if (linkedItem.Type.Name.Equals("Bug", StringComparison.OrdinalIgnoreCase))
                    {
                        if (linkedItem.State == "Done")
                            bugMetrics["Fixed"].Add(linkedItem.Id);
                        else if (linkedItem.State != "Removed")
                            bugMetrics["Deferred"].Add(linkedItem.Id);
                    }
                }
            }

            bool testcasefound = false;
            foreach (ITestSuiteEntry tse in itsb.TestCases)
            {
                if (tse.EntryType == TestSuiteEntryType.TestCase)
                {
                    if (tse.TestCase != null)
                    {

                        testCases.Add(tse.TestCase);
                        testcasefound = true;
                    }
                }
            }
            if (testcasefound == true)
            {
                if (sheetno > defaultSheets)
                {
                    xlBook.Sheets.Add(Type.Missing, xlBook.Sheets[sheetno - 1], Type.Missing, Type.Missing);
                }
                xlWorkSheet = (Excel.Worksheet)xlBook.Worksheets.get_Item(sheetno);
                Sname = formatsheetname(itsb.Title);
                sheetcount = xlBook.Worksheets.Count;
                for (int k = 1; k <= sheetcount; k++)
                {
                    xlWSheet = (Excel.Worksheet)xlBook.Worksheets.get_Item(k);
                    if (Sname == xlWSheet.Name)
                    {
                        if (itsb.Parent.Title.Length > 15)
                        {
                            int parentsuitenamelength = itsb.Parent.Title.Length;
                            Sname = string.Concat(formatsheetname(itsb.Parent.Title.Substring((parentsuitenamelength - 15), 15)), "_", Sname);
                        }
                        else
                        {
                            Sname = string.Concat(formatsheetname(itsb.Parent.Title), "_", Sname);
                        }

                    }
                }
                if (Sname.Length > 30)
                { Sname = Sname.Substring(0, 30); }
                xlWorkSheet.Name = Sname;

                export(xlWorkSheet, testCases);
                sheetno++;
                testcasefound = false;
                testCases.Clear();
            }


            if (itsb.TestSuiteType == TestSuiteType.StaticTestSuite)
            {
                IStaticTestSuite staticsuite = itsb as IStaticTestSuite;
                foreach (ITestSuiteBase tse in staticsuite.SubSuites)
                {
                    exportmultisheet(tse, xlBook);
                }
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            Stopwatch mainStopWatch = new Stopwatch();
            mainStopWatch.Start();
            unfixedBugs = "";

            if (txtFileName.Text != null && txtFileName.Text != "" && treeView_suite.SelectedNode != null && txtSaveFolder.Text != null && txtSaveFolder.Text != "")
            {
                this.Cursor = Cursors.WaitCursor;
                btnExport.Enabled = false;
                btnCancel.Enabled = false;
                btnHelp.Enabled = false;
                btnTeamProject.Enabled = false;
                btnFolderBrowse.Enabled = false;
                comBoxTestPlan.Enabled = false;
                int k;
                xlApp = new Excel.Application();

                k = Convert.ToInt32(treeView_suite.SelectedNode.Name.ToString());
                ITestSuiteBase suite1 = _teamProject.TestSuites.Find(k);
                Get_TestCases(suite1);
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                sheetno = 1;
                defaultSheets = xlWorkBook.Sheets.Count;
                if (SeparateSheets.Checked == true)
                {
                    if (suite1.TestSuiteType.ToString() == "StaticTestSuite")
                    {
                        testCases.Clear();
                        resetMetrics();
                        exportmultisheet(suite1, xlWorkBook);
                    }
                }
                else
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(sheetno);
                    xlWorkSheet.Name = formatsheetname(suite1.Title);
                    if (testCases.Count > 0)
                    {
                        resetMetrics();
                        export(xlWorkSheet, testCases);
                        testCases.Clear();
                    }
                }

                createTestReportDetailsSheet();

                // save excel file
                try
                {
                    xlWorkBook.SaveAs(txtSaveFolder.Text + "\\" + txtFileName.Text + ".xlsx", Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();

                    releaseObject(xlApp);
                    releaseObject(xlWorkBook);
                    releaseObject(xlWorkSheet);
                    this.Cursor = Cursors.Arrow;
                    btnExport.Enabled = true;
                    btnCancel.Enabled = true;
                    btnHelp.Enabled = true;
                    btnTeamProject.Enabled = true;
                    btnFolderBrowse.Enabled = true;
                    comBoxTestPlan.Enabled = true;
                    txtFileName.Text = "";
                    mainStopWatch.Stop();
                    MessageBox.Show("Test Cases exported successfully to specified file. " + mainStopWatch.Elapsed.ToString(), "Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                    Console.WriteLine("Main Time Elapsed={0}", mainStopWatch.Elapsed);
                }
                catch (Exception ex)
                {
                    if (ex.Message == "Cannot access '" + txtFileName.Text + ".xls'.")
                    {
                        MessageBox.Show("File with same name exists in specified location", "File Exists", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        txtFileName.Text = "";
                    }
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
            else
            {
                MessageBox.Show("All fields are not populated.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void createTestReportDetailsSheet()
        {
            Excel.Worksheet sheet = (Excel.Worksheet)xlWorkBook.Sheets.Add(xlWorkBook.Sheets[1], Type.Missing, Type.Missing, Type.Missing);
            sheet.Name = "Test Report Details!";

            sheet.Cells.Font.Name = "Calibri";
            sheet.Cells.Font.Size = 11;

            sheet.Columns["A"].ColumnWidth = 34;
            sheet.Columns["B"].ColumnWidth = 60;
            sheet.Columns["C"].ColumnWidth = 11;
            sheet.Columns["D"].ColumnWidth = 16;
            sheet.Columns["E"].ColumnWidth = 41;

            Excel.Range colB = (Excel.Range)sheet.Columns["B"];
            Excel.Range colE = (Excel.Range)sheet.Columns["E"];
            colB.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            colE.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            // Test Report Details section
            var titleCell = (Excel.Range)sheet.Cells[1, 1];
            titleCell.Value2 = "Test Report Details";
            titleCell.Font.Bold = true;
            titleCell.Font.Italic = true;
            titleCell.Font.Underline = true;
            titleCell.Font.Size = 12;

            var detailsLabels = new[]
            {
            Tuple.Create(3, "Product:"),
            Tuple.Create(4, "Tested release:"),
            Tuple.Create(5, "Date:"),
            Tuple.Create(6, "Test plan:"),
            Tuple.Create(7, "Tested by:"),
            Tuple.Create(8, "Test environment:")
            };

            foreach (var lbl in detailsLabels)
            {
                var cell = (Excel.Range)sheet.Cells[lbl.Item1, 1];
                cell.Value2 = lbl.Item2;
                cell.Font.Bold = true;
            }

            // Approval Section
            var approvalCell = (Excel.Range)sheet.Cells[1, 4];
            approvalCell.Value2 = "Approval";
            approvalCell.Font.Bold = true;
            approvalCell.Font.Italic = true;
            approvalCell.Font.Underline = true;
            approvalCell.Font.Size = 12;

            var approvalLabels = new[]
            {
            Tuple.Create(3, "Approved by:"),
            Tuple.Create(4, "Date:")
            };
            foreach (var lbl in approvalLabels)
            {
                var cell = (Excel.Range)sheet.Cells[lbl.Item1, 4];
                cell.Value2 = lbl.Item2;
                cell.Font.Bold = true;
                cell.Font.Size = 11;
            }

            // Summary section
            var summaryCell = (Excel.Range)sheet.Cells[10, 1];
            summaryCell.Value2 = "Test Report Summary";
            summaryCell.Font.Bold = true;
            summaryCell.Font.Italic = true;
            summaryCell.Font.Underline = true;
            summaryCell.Font.Size = 12;

            var summaryLabels = new[] {
                Tuple.Create(12, "Functional and regression test cases"),
                Tuple.Create(13, "Total:"),
                Tuple.Create(14, "Executed:"),
                Tuple.Create(15, "Passed:"),
                Tuple.Create(16, "Failed:"),
                
                Tuple.Create(18, "Integration and Unit tests"),
                Tuple.Create(19, "Total:"),
                Tuple.Create(20, "Executed:"),
                Tuple.Create(21, "Passed:"),
                Tuple.Create(22, "Failed:"),
                
                Tuple.Create(24, "Bugs"),
                Tuple.Create(25, "Identified:"),
                Tuple.Create(26, "Fixed:"),
                Tuple.Create(27, "Deffered:"),
                Tuple.Create(28, "List of not fixed bugs:")
            };

            foreach (var lbl in summaryLabels)
            {
                var cell = (Excel.Range)sheet.Cells[lbl.Item1, 1];
                cell.Value2 = lbl.Item2;
                if(new[] { 12, 18, 24, 28}.Contains(lbl.Item1))
                    cell.Font.Bold = true;
            }

            // Add metrics values
            var metrics = new[]
            {
                Tuple.Create(13 , testMetrics["Total"]),
                Tuple.Create(14 , testMetrics["Executed"]),
                Tuple.Create(15 , testMetrics["Passed"]),
                Tuple.Create(16 , testMetrics["Failed"]),

                Tuple.Create(25 , bugMetrics["Fixed"].Count() + bugMetrics["Deferred"].Count()),
                Tuple.Create(26 , bugMetrics["Fixed"].Count()),
                Tuple.Create(27 , bugMetrics["Deferred"].Count())
            };
            
            foreach (var mt in metrics)
            {
                var cell = (Excel.Range)sheet.Cells[mt.Item1, 2];
                cell.Value2 = mt.Item2;
            }

            sheet.Cells[28, 2] = string.Join(", ", bugMetrics["Deferred"]);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnAbout_Click(object sender, EventArgs e)
        {
            FrmHelp frmAbout = new FrmHelp();
            frmAbout.ShowDialog();
        }

        private void FrmMain_Load(object sender, EventArgs e)
        {

        }

        private void txtTeamProject_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void lblWelcomeMessage_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void lblTestPlan_Click(object sender, EventArgs e)
        {

        }

    }
}
