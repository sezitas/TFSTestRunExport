using System;
using System.Drawing;
//using System.Windows;
using System.Windows.Forms;

using System.Text.RegularExpressions;
//using System.Windows.Documents;
//using System.Windows.Navigation;
//using System.Windows.Shapes;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

//comment

namespace TestCaseExport
{
    public partial class FrmMain : Form
    {
        private TfsTeamProjectCollection _tfs;
        private ITestManagementTeamProject _teamProject;
        private ITestPlanCollection plans;
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
        //string upperBound = "a";
        //string lowerBound = "a";
        int upperBound = 0;
        int lowerBound = 0;
        int sheetno = 1;
        int defaultSheets;
        int sheetcount;
        string Sname;


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

            //text = text.Replace("<HTML><BODY><P><SPAN>", "");
            //text = text.Replace("</SPAN></P></BODY></HTML>", "");
            //text = text.Replace("<HTML><BODY><P /></BODY></HTML>", "");
            //text = text.Replace("<HTML><BODY><P>", "");
            //text = text.Replace("</P></BODY></HTML>", "");
            //text = text.Replace("</SPAN></P><P>", "");
            //text = text.Replace("</P><P /></BODY></HTML>", "");
            //text = text.Replace("</P><P /><P>", "");
            //text = text.Replace("<SPAN>", "");
            //text = text.Replace("</SPAN>", "");
            //text = text.Replace("</P><P><SPAN/>", "");
            //text = text.Replace("</P><P>", "");
            //text = text.Replace("<dir><font size=3 face=Calibri><font size=3 face=Calibri><span lang=EN>", "");
            //text = text.Replace("<p dir=ltr align=left>", "");
            //text = text.Replace("<span lang=EN-GB>", "");
            //text = text.Replace("<p dir=ltr align=left>&#160;</p>", "");
            //text = text.Replace("<p dir=ltr align=left>", "");
            //text = text.Replace("<dir><font size=3 face=Calibri><font size=3 face=Calibri><span lang=EN>", "");
            //text = text.Replace("</p>", "");
            //text = text.Replace("</font></font></b>", "");
            //text = text.Replace("</font></font></font>", "");
            //text = text.Replace("</i></b><span style=\">", "");
            //text = text.Replace("&#160;", "");
            //text = text.Replace("<strong><font size=3><font face=Calibri>", "");
            //text = text.Replace("</font></font>", "");
            //text = text.Replace("</strong>", "");
            //text = text.Replace("</i><span style=\">", "");
            //text = text.Replace("</i>", "");
            //text = text.Replace("<DIV><P /></DIV>", "");
            //text = text.Replace("</P><P /><P /></DIV>", "");
            //text = text.Replace("</P><P /></DIV>", "");       
            //text = text.Replace("<DIV><P>", "");
            //text = text.Replace("</P></DIV>", "");
            //text = text.Replace("</SPAN>", "");
            //text = text.Replace("<SPAN>", "");
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


        private void export(Excel.Worksheet xlWorkSheet, ITestCaseCollection testcases)
        {
            xlWorkSheet.Cells[1, 1] = "TC No";
            xlWorkSheet.Cells[1, 2] = "Test Case Title";
            xlWorkSheet.Cells[1, 3] = "Summary";
            xlWorkSheet.Cells[1, 4] = "Action";
            xlWorkSheet.Cells[1, 5] = "Expected Result";
            //xlWorkSheet.Cells[1, 6] = "Actual Result";
            xlWorkSheet.Cells[1, 6] = "Pass/Fail";
            xlWorkSheet.Cells[1, 7] = "Bug ID";
            xlWorkSheet.Cells[1, 8] = "Comments";

            (xlWorkSheet.Columns["A", Type.Missing]).ColumnWidth = 9;
            (xlWorkSheet.Columns["B", Type.Missing]).ColumnWidth = 35;
            (xlWorkSheet.Columns["C", Type.Missing]).ColumnWidth = 30;
            (xlWorkSheet.Columns["D", Type.Missing]).ColumnWidth = 50;
            (xlWorkSheet.Columns["E", Type.Missing]).ColumnWidth = 50;
            //(xlWorkSheet.Columns["F", Type.Missing]).ColumnWidth = 30;
            (xlWorkSheet.Columns["F", Type.Missing]).ColumnWidth = 12;
            (xlWorkSheet.Columns["G", Type.Missing]).ColumnWidth = 12;
            (xlWorkSheet.Columns["H", Type.Missing]).ColumnWidth = 20;

            foreach (ITestCase testCase in testCases)
            {

                //upperBound = "a";
                //lowerBound = "a";
                //xlWorkSheet.Cells[row, col] = testCase.Title;

                upperBound = row;
                TestActionCollection testActions = testCase.Actions;
                var testResults = _teamProject.TestResults.ByTestId(testCase.Id);

                int i = 1;
                if (testActions.Count == 0)
                {
                    xlWorkSheet.Cells[row, 1] = testCase.Id.ToString();
                    row++;
                }
                else
                {
                    foreach (ITestAction action in testActions)
                    {
                        ISharedStep shared_step = null;
                        ISharedStepReference shared_ref = action as ISharedStepReference;
                        if (shared_ref != null)
                        {
                            shared_step = shared_ref.FindSharedStep();
                            foreach (ITestAction shr_action in shared_step.Actions)
                            {
                                var stest_step = shr_action as ITestStep;
                                xlWorkSheet.Cells[row, 4] = removehtmltags(stest_step.Title.ToString());
                                xlWorkSheet.Cells[row, 5] = removehtmltags(stest_step.ExpectedResult.ToString());
                                xlWorkSheet.Cells[row, 1] = testCase.Id.ToString() + "." + i;
                                row++;
                                i++;
                            }

                        }
                        else
                        {
                            var testStep = action as ITestStep;
                            xlWorkSheet.Cells[row, 4] = removehtmltags(testStep.Title.ToString());
                            xlWorkSheet.Cells[row, 5] = removehtmltags(testStep.ExpectedResult.ToString());
                            xlWorkSheet.Cells[row, 1] = testCase.Id.ToString() + "." + i;
                            row++;
                            i++;

                        }
                        if (ExportResults.Checked == true)
                        {
                            foreach (ITestCaseResult result in testResults)
                            {
                                int top = result.Iterations.Count;
                                if (top > 0)
                                {
                                    var topIteration = result.Iterations[top];
                                    if (topIteration == null)
                                        continue;
                                    int actionindex = testActions.IndexOf(action.Id);
                                    if (actionindex < topIteration.Actions.Count)
                                    {
                                        var actionResult = topIteration.Actions[actionindex];
                                        if (actionResult.Outcome.ToString() != "None" && actionResult.Outcome.ToString() != "Unspecified")
                                        {
                                            xlWorkSheet.Cells[(row - 1), 6] = actionResult.Outcome.ToString();
                                            xlWorkSheet.Cells[(row - 1), 8] = actionResult.ErrorMessage.ToString();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                lowerBound = (row - 1);

                xlWorkSheet.get_Range("c" + upperBound, "c" + lowerBound).Merge(false);
                chartRange = xlWorkSheet.get_Range("c" + upperBound, "c" + lowerBound);
                chartRange.FormulaR1C1 = removehtmltags(testCase.Description.ToString());

                xlWorkSheet.get_Range("b" + upperBound, "b" + lowerBound).Merge(false);
                chartRange = xlWorkSheet.get_Range("b" + upperBound, "b" + lowerBound);
                chartRange.FormulaR1C1 = testCase.Title.ToString();

                chartRange.HorizontalAlignment = 1;
                chartRange.VerticalAlignment = 1;

                //populate bugs

                if (ExportResults.Checked == true)
                {
                    Query query = new Query(_store, string.Format("SELECT [Target].[System.Id] FROM WorkItemLinks WHERE ([Source].[System.Id] = {0}) and ([Source].[System.WorkItemType] = 'Test Case')  And ([Target].[System.WorkItemType] = 'Bug')mode(MustContain)", testCase.Id));

                    WorkItemLinkInfo[] workItemLinkInfoArray = null;
                    if (query.IsLinkQuery)
                    {

                        workItemLinkInfoArray = query.RunLinkQuery();

                    }

                    else
                    {

                        throw new Exception("Run link query fail. Query passed is not a link query");

                    }
                    string bug_list = "";
                    bool multibug = false;
                    for (int k = 0; k < workItemLinkInfoArray.Length; k++)
                    {
                        if (workItemLinkInfoArray[k].LinkTypeId != 0)
                        {
                            if (multibug == true)
                            {
                                bug_list = bug_list + ", ";
                            }
                            bug_list = bug_list + workItemLinkInfoArray[k].TargetId.ToString();
                            multibug = true;
                        }
                    }


                    xlWorkSheet.get_Range("G" + upperBound, "G" + lowerBound).Merge(false);
                    chartRange = xlWorkSheet.get_Range("G" + upperBound, "G" + lowerBound);
                    chartRange.FormulaR1C1 = bug_list;
                }


            }
            lowerBound = (row - 1);
            chartRange = xlWorkSheet.get_Range("H" + lowerBound, "H1");
            //chartRange.Font.Bold = true;
            //chartRange.Interior.Color = 18018018;


            chartRange = xlWorkSheet.get_Range("a1", "H" + lowerBound);
            chartRange.Cells.WrapText = true;
            chartRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;


            //chartRange.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

            chartRange = xlWorkSheet.get_Range("a1", "H1");
            chartRange.Borders.LineStyle = Excel.XlLineStyle.xlDouble;
            chartRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            chartRange.Font.Bold = true;

            row = 2;
            upperBound = 0;
            lowerBound = 0;

        }

        private void exportmultisheet(ITestSuiteBase itsb, Excel.Workbook xlBook)
        {


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
                        exportmultisheet(suite1, xlWorkBook);

                    }
                }
                else
                {
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(sheetno);
                    xlWorkSheet.Name = formatsheetname(suite1.Title);
                    if (testCases.Count > 0)
                    {
                        export(xlWorkSheet, testCases);
                        testCases.Clear();
                    }


                }



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
                    MessageBox.Show("Test Cases exported successfully to specified file.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

                    //txtTeamProject.Text = "";
                    //comBoxTestPlan.Items.Clear();

                    //txtSaveFolder.Text = "";

                }
                catch (Exception ex)
                {
                    if (ex.Message == "Cannot access '" + txtFileName.Text + ".xls'.")
                    {
                        MessageBox.Show("File with same name exists in specified location", "File Exists", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        txtFileName.Text = "";
                    }
                    //else
                    //{
                    //MessageBox.Show("Application has encountered Fatal Errro. \nPlease contact your System Administrator.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //}
                }

            }
            else
            {
                MessageBox.Show("All fields are not populated.", "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
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
