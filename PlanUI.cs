using Microsoft.TeamFoundation.TestManagement.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace TestPlanViewer
{
    /// <summary>
    /// Main UI class
    /// </summary>
    public partial class PlanUI : Form
    {

        /// <summary>
        /// Connection to the test manager results
        /// </summary>
        private TFSConnection connection;

        /// <summary>
        /// Current sort order
        /// </summary>
        private SortOrder currentSort;

        /// <summary>
        /// Column to sort on
        /// </summary>
        private int currentSortColumn;

        /// <summary>
        /// List for storing executing threads
        /// </summary>
        private List<Thread> childThreads;

        /// <summary>
        /// Initializes a new instance of the TestResultUI class
        /// </summary>
        public PlanUI()
        {
            this.InitializeComponent();
            this.currentSort = SortOrder.Ascending;
            this.currentSortColumn = 0;
            this.childThreads = new List<Thread>();
        }

        /// <summary>
        /// Query the server for results
        /// </summary>
        /// <param name="sender">The system send object</param>
        /// <param name="e">The system event</param>
        private void GetResultsButton_Click(object sender, EventArgs e)
        {
            bool getResultsEnabled = this.GetResultsButton.Enabled;
            bool switchEnabled = this.SwitchPlanButton.Enabled;
            bool saveEnabled = this.SaveButton.Enabled;
            this.GetResultsButton.Enabled = false;
            this.SwitchPlanButton.Enabled = false;
            this.SaveButton.Enabled = false;

            this.NewQuery(false, getResultsEnabled, switchEnabled, saveEnabled);
        }

        /// <summary>
        /// The switch plan button has been clicked
        /// </summary>
        /// <param name="sender">The system send object</param>
        /// <param name="e">The system event</param>
        private void SwitchPlanButton_Click(object sender, EventArgs e)
        {
            bool getResultsEnabled = this.GetResultsButton.Enabled;
            bool switchEnabled = this.SwitchPlanButton.Enabled;
            bool saveEnabled = this.SaveButton.Enabled;
            this.GetResultsButton.Enabled = false;
            this.SwitchPlanButton.Enabled = false;
            this.SaveButton.Enabled = false;

            this.NewQuery(true, getResultsEnabled, switchEnabled, saveEnabled);
        }

        /// <summary>
        /// The save button has been clicked
        /// </summary>
        /// <param name="sender">The system send object</param>
        /// <param name="e">The system event</param>
        private void SaveButton_Click(object sender, EventArgs e)
        {
            List<TestResult> results = new List<TestResult>();
            foreach (TreeNode node in this.StructureView.Nodes)
            {
                this.GetResultList(results, node);
            }

            ExportToExcel doc = new ExportToExcel();
            doc.LaunchExcel(new AllResults(results));
        }

        /// <summary>
        /// The user has chosen to look at a test suite or single test case
        /// </summary>
        /// <param name="sender">The system send object</param>
        /// <param name="e">The system event</param>
        private void StructureView_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            // Make sure we have a node
            TreeNode node = e.Node;
            if (node == null)
            {
                return;
            }

            // Don't load the data unless we are left clicking on the folder
            TreeViewHitTestLocations location = this.StructureView.HitTest(e.Location).Location;
            if ((location != TreeViewHitTestLocations.Image && location != TreeViewHitTestLocations.Label) || e.Button != MouseButtons.Left)
            {
                return;
            }

            // Make sure we have result structure
            TestResult result = (TestResult)node.Tag;
            if (result == null)
            {
                return;
            }

            // Disable the controls so the user sees something happening
            this.ResultsListView.Enabled = false;
            this.StructureView.Enabled = false;

            // Clear the items in the listview
            this.ResultsListView.Items.Clear();

            // Select the node
            ((TreeView)sender).SelectedNode = node;
            if (result.TestID >= 0)
            {
                // Add results for a single test
                this.AddTestResult(result);
                this.SelectedCounts(result);
            }
            else
            {
                // Add a suite and it's sub suites
                this.AddChildTestResults(node);
                this.SelectedCounts(node);
            }

            // Enable the controls
            this.ResultsListView.Enabled = true;
            this.StructureView.Enabled = true;
        }

        /// <summary>
        /// The user double clicked a specific node in the tree view
        /// </summary>
        /// <param name="sender">The system send object</param>
        /// <param name="e">The system event</param>
        private void StructureView_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            // Only do left double clicks
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            // Make sure we have a node
            TreeNode node = e.Node;
            if (node == null)
            {
                return;
            }

            // Make sure we have result structure
            TestResult result = (TestResult)node.Tag;
            if (result == null || result.TestID < 0)
            {
                return;
            }

            this.OpenWebTestCase(result.TestID, result.SuiteID);
        }

        /// <summary>
        /// Add child results for the given parent tree node
        /// </summary>
        /// <param name="node">The parent node</param>
        private void AddChildTestResults(TreeNode node)
        {
            // Look at all the child nodes
            foreach (TreeNode childNode in node.Nodes)
            {
                TestResult result = (TestResult)childNode.Tag;
                if (result.TestID >= 0)
                {
                    this.AddTestResult(result);
                }
                else
                {
                    this.AddChildTestResults(childNode);
                }
            }
        }

        /// <summary>
        /// Add child results for the given parent tree node
        /// </summary>
        /// <param name="node">The parent node</param>
        /// <param name="results">A list of test results to store found results</param>
        private void GetChildTestResults(TreeNode node, List<TestResult> results)
        {
            // Look at all the child nodes
            foreach (TreeNode childNode in node.Nodes)
            {
                TestResult result = (TestResult)childNode.Tag;
                if (result.TestID >= 0)
                {
                    results.Add(result);
                }
                else
                {
                    this.GetChildTestResults(childNode, results);
                }
            }
        }

        /// <summary>
        /// A column header is clicked so sort the results
        /// </summary>
        /// <param name="sender">The system send object</param>
        /// <param name="e">The system event</param>
        private void ResultsListView_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == this.currentSortColumn)
            {
                // Reverse the current sort direction for this column.
                if (this.currentSort == SortOrder.Ascending)
                {
                    this.currentSort = SortOrder.Descending;
                }
                else
                {
                    this.currentSort = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                this.currentSortColumn = e.Column;
                this.currentSort = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            this.ResultsListView.ListViewItemSorter = new ListViewItemComparer(this.currentSortColumn, this.currentSort);
            this.ResultsListView.Sort();
        }

        /// <summary>
        /// Copy selected results to the clipboard
        /// </summary>
        /// <param name="sender">The system send object</param>
        /// <param name="e">The system event</param>
        private void ResultsListView_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.C && e.Control)
                {
                    StringBuilder buffer = new StringBuilder();

                    // Get the column headers
                    for (int i = 0; i < this.ResultsListView.Columns.Count; i++)
                    {
                        buffer.Append(this.ResultsListView.Columns[i].Text);
                        buffer.Append("\t");
                    }

                    buffer.Append("\n");

                    // Loop over all the selected rows
                    foreach (ListViewItem item in this.ResultsListView.SelectedItems)
                    {
                        for (int i = 0; i < this.ResultsListView.Columns.Count; i++)
                        {
                            buffer.Append(item.SubItems[i].Text);
                            buffer.Append("\t");
                        }

                        buffer.Append("\n");
                    }

                    Clipboard.SetText(buffer.ToString());
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Failed to copy results to clipboard because:\r\n" + exception.Message, "Copy To Clipboard Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Open a specific test case in the web version of TFS
        /// </summary>
        /// <param name="sender">The system send object</param>
        /// <param name="e">The system event</param>
        private void ResultsListView_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Left)
                {
                    // Get the test ID
                    ListViewHitTestInfo item = this.ResultsListView.HitTest(e.X, e.Y);
                    ListViewItem.ListViewSubItem sub = item.Item.SubItems[2];
                    int outTest;



                    if (int.TryParse(sub.Text, out outTest))
                    {
                        this.OpenWebTestCase(outTest, ((int)item.Item.Tag));
                    }
                    else
                    {
                        // Open the test case
                        MessageBox.Show("Could not open test case because:\r\nThe ID '" + sub.Text + "' is not numeric.", "Open Test Case Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Failed open test case because:\r\n" + exception.Message, "Open Test Case Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Do a prartial update of the results
        /// </summary>
        /// <param name="sender">The system send object</param>
        /// <param name="e">The system event</param>
        private void GetResultsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ContextMenuStrip context = (ContextMenuStrip)((ToolStripMenuItem)sender).Owner;
            TreeView structureView = (TreeView)context.SourceControl;
            TreeNode node = structureView.GetNodeAt(structureView.PointToClient(context.Location));

            // No node selected
            if (node == null)
            {
                return;
            }

            bool alreadyBusy = false;

            // Make sure the busy icon is visible
            if (this.ToolStripStatusBusy.Visible == true)
            {
                alreadyBusy = true;
            }
            else
            {
                this.ToolStripStatusBusy.Visible = true;
            }

            // Disable the views so the user doesn't think they can use it
            this.StructureView.Enabled = false;
            this.ResultsListView.Enabled = false;

            this.connection.UpdateWithLatestResult(node);
            
            // Clear the listview and than add the updated values
            this.ResultsListView.Items.Clear();
            this.AddChildTestResults(node);
            this.NewCounts();

            // Enable the views
            structureView.Enabled = true;
            this.ResultsListView.Enabled = true;

            // Only get rid of the image if we set it to display in the first place
            if (alreadyBusy)
            {
                this.ToolStripStatusBusy.Visible = true;
            }
            else
            {
                this.ToolStripStatusBusy.Visible = false;
            }
        }

        /// <summary>
        /// Export specific data from a specific node
        /// </summary>
        /// <param name="sender">The system send object</param>
        /// <param name="e">The system event</param>
        private void ExportToExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Find the selected node
            ContextMenuStrip context = (ContextMenuStrip)((ToolStripMenuItem)sender).Owner;
            TreeView structureView = (TreeView)context.SourceControl;
            TreeNode node = structureView.GetNodeAt(structureView.PointToClient(context.Location));

            if (node == null)
            {
                // No node selected so quite
                return;
            }

            // What should we save this file as
            string file = SharedUtils.GetFileName("Excel Open XML | *.xlsm", "xlsm");
            if (string.IsNullOrEmpty(file))
            {
                return;
            }

            // Get the results
            List<TestResult> results = new List<TestResult>();
            this.GetChildTestResults(node, results);

            // Create the excel file
            ExportToExcel doc = new ExportToExcel();
            doc.LaunchExcel(new AllResults(results), file);
        }

        /// <summary>
        /// The form is closing
        /// </summary>
        /// <param name="sender">The system send object</param>
        /// <param name="e">The system event</param>
        private void TestResultUI_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Make sure we don't leave threads spinning
            this.KillQueryThreads();
        }

        /// <summary>
        /// Query the server for updated data
        /// </summary>
        /// <param name="switchPlan">Are we switching test suites</param>
        /// <param name="getResultsWasEnabled">Was the get results button enable before the query</param>
        /// <param name="switchWasEnabled">Was the switch results button enable before the query</param>
        /// <param name="saveWasEnabled">Was the save button enable before the query</param>
        private void NewQuery(bool switchPlan, bool getResultsWasEnabled, bool switchWasEnabled, bool saveWasEnabled)
        {
            try
            {
                // If we are switching test plans or the tree view is empty
                if (switchPlan || this.StructureView.Nodes.Count == 0)
                {
                    // Setup a connection if doesn't already exist
                    if (this.connection == null || switchPlan)
                    {
                        // Try to get a new connection
                        TFSConnection temp = new TFSConnection(this.StructureView);

                        if (temp == null || string.IsNullOrEmpty(temp.GetTestPlanName()))
                        {
                            // No plan was selected so bail
                            return;
                        }

                        this.connection = temp;
                    }

                    // Not authenticated so stop the query
                    if (!this.connection.AuthenticationCheck())
                    {
                        return;
                    }

                    // Select a test plan
                    string plan = this.connection.GetTestPlanName();

                    // Make sure all existing queries are stopped and the tree view is cleared
                    this.KillQueryThreads();
                    this.StructureView.Nodes.Clear();
                    this.ResultsListView.Items.Clear();

                    // Update the UI
                    this.GetResultsButton.Text = "Refresh";
                    this.GetResultsButton.Enabled = false;
                    this.ToolStripStatusBusy.Visible = true;
                    this.SwitchPlanButton.Enabled = true;
                    this.NewCounts();

                    // Get the temp cache file setup
                    string fileName = this.NameSpecificFileName(plan);
                    string tempPath = Path.Combine(Path.GetTempPath(), fileName);
                    string baseCachedFile = Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "StartingPoints", fileName);

                    // If we have a cached copy us it
                    if (!File.Exists(tempPath))
                    {
                        // No cached copy found so use our base cached files
                        if (File.Exists(baseCachedFile))
                        {
                            FileInfo baseFile = new FileInfo(baseCachedFile);
                            baseFile.CopyTo(tempPath, true);
                            FileInfo cachedFile = new FileInfo(tempPath);
                            cachedFile.LastAccessTimeUtc = baseFile.LastAccessTimeUtc;
                        }
                    }
                    else if (File.Exists(baseCachedFile))
                    {
                        // Check if the base cache is new than the temp cache
                        FileInfo baseFile = new FileInfo(baseCachedFile);
                        FileInfo cachedFile = new FileInfo(tempPath);

                        if (cachedFile.LastAccessTimeUtc < baseFile.LastAccessTimeUtc)
                        {
                            baseFile.CopyTo(tempPath, true);
                            cachedFile.LastAccessTimeUtc = baseFile.LastAccessTimeUtc;
                        }
                    }

                    // Update the current data and kick off a worker thread to get the data from the server
                    Thread thread;
                    if (this.LoadTree(tempPath))
                    {
                        this.ToolStripStatusLabel.Text = "Updating test results for " + plan;
                        this.ToolStripStatusLastUpdate.Text = (string)this.StructureView.Tag;
                        thread = new Thread(this.FullViewUpdate);
                    }
                    else
                    {
                        this.ToolStripStatusLastUpdate.Text = "Never";
                        this.ToolStripStatusLabel.Text = "Getting plan structure for " + plan;
                        thread = new Thread(this.StructFirstViewUpdate);
                    }

                    // Kick off the thread
                    this.childThreads.Add(thread);
                    thread.Start();
                }
                else
                {
                    // Update the UI and kick off the thread
                    string appendText = string.Empty;
                    if (this.StructureView.Nodes != null && this.StructureView.Nodes.Count > 0)
                    {
                        appendText = "for " + this.StructureView.Nodes[0].Text;
                    }

                    this.ToolStripStatusLabel.Text = "Getting test results " + appendText;
                    this.GetResultsButton.Enabled = false;
                    this.ToolStripStatusBusy.Visible = true;
                    Thread thread = new Thread(this.FullViewUpdate);
                    this.childThreads.Add(thread);
                    thread.Start();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(string.Format("Failed to query data because:\r\n{0}", e.Message), "Query Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                // If a button is disable, but was enabled before query enable it
                this.GetResultsButton.Enabled = this.GetResultsButton.Enabled == true ? true : getResultsWasEnabled;
                this.SwitchPlanButton.Enabled = this.SwitchPlanButton.Enabled == true ? true : switchWasEnabled;
                this.SaveButton.Enabled = this.SaveButton.Enabled == true ? true : saveWasEnabled;
            }
        }

        /// <summary>
        /// Update the TreeView scructure and than get results
        /// </summary>
        private void StructFirstViewUpdate()
        {
            // Get the structure
            this.ViewUpdate(false);

            // Get the plan name
            string appendText = string.Empty;
            if (this.StructureView.Nodes != null && this.StructureView.Nodes.Count > 0)
            {
                appendText = "for " + this.StructureView.Nodes[0].Text;
            }

            // Update the UI message
            Action updateMessageAction = () => this.ToolStripStatusLabel.Text = "Getting test results " + appendText;
            this.Invoke(updateMessageAction);

            // Get the detailed data
            this.ViewUpdate(true);
        }

        /// <summary>
        /// Do a full update, structure and results
        /// </summary>
        private void FullViewUpdate()
        {
            this.ViewUpdate(true);
        }

        /// <summary>
        /// Update the TreeView
        /// </summary>
        /// <param name="full">True if you are updating the structure and results. False if you are only doing the structure</param>
        private void ViewUpdate(bool full)
        {
            // Create a swap tree view to do all the work so the UI is not affected until the data is collected
            TreeView swapTree = new TreeView();
            foreach (TreeNode node in this.StructureView.Nodes)
            {
                swapTree.Nodes.Add((TreeNode)node.Clone());
            }

            this.connection.View = swapTree;
            this.connection.GetSuiteStructure(full);

            // Update the UI
            Action finishAction = () => this.FinishViewUpdate(swapTree, full);
            this.Invoke(finishAction);
        }

        /// <summary>
        /// Once a swap TreeView has been updated swap it out for the current and update the UI
        /// </summary>
        /// <param name="swapView">The TreeView to swap in</param>
        /// <param name="unlock">Should this step unlock the UI</param>
        private void FinishViewUpdate(TreeView swapView, bool unlock)
        {
            // Make sure the icons are set
            foreach (TreeNode node in swapView.Nodes)
            {
                this.SetIcons(node);
            }

            TreeNode[] rootNodeArray = new TreeNode[swapView.Nodes.Count];
            swapView.Nodes.CopyTo(rootNodeArray, 0);
            this.StructureView.Nodes.Clear();
            this.StructureView.Nodes.AddRange(rootNodeArray);
            this.StructureView.Sort();
            this.ToolStripStatusLastUpdate.Text = DateTime.Now.ToString("f");
            this.StructureView.Tag = this.ToolStripStatusLastUpdate.Text;
            this.StructureView.Refresh();
            this.ResultsListView.Items.Clear();
            this.NewCounts();

            // Unlocking so enble buttons
            if (unlock)
            {
                this.GetResultsButton.Enabled = true;
                this.ToolStripStatusLabel.Text = string.Empty;
                this.ToolStripStatusBusy.Visible = false;
                this.SaveButton.Enabled = true;

                // Clear selected counts
                this.PassedSelectedLabel.Text = "0";
                this.FailedSelectedLabel.Text = "0";
                this.BlockedSelectedLabel.Text = "0";
                this.QueuedSelectedLabel.Text = "0";
                this.InconclusiveSelectedLabel.Text = "0";
                this.NotExecutedSelectedLabel.Text = "0";
                this.ActiveSelectedLabel.Text = "0";
                this.OtherSelectedLabel.Text = "0";
                this.TotalSelectedLabel.Text = "0";
            }

            // Cache the results
            if (swapView.Nodes != null && swapView.Nodes.Count > 0)
            {
                this.SaveTree(Path.Combine(Path.GetTempPath(), this.NameSpecificFileName(swapView.Nodes[0].Text)));
            }
        }

        /// <summary>
        /// Update the UI counts
        /// </summary>
        private void NewCounts()
        {
            // Clear old values
            this.PassedTotal.Text = "0";
            this.FailedTotal.Text = "0";
            this.BlockedTotal.Text = "0";
            this.QueuedTotal.Text = "0";
            this.InconclusiveTotal.Text = "0";
            this.OtherTotal.Text = "0";
            this.ActiveTotal.Text = "0";

            List<TestResult> results = new List<TestResult>();
            foreach (TreeNode node in this.StructureView.Nodes)
            {
                this.GetResultList(results, node);
            }

            AllResults cumulativeResults = new AllResults(results);

            this.PassedTotal.Text = cumulativeResults.GetOutcomeCount(TestOutcome.Passed, false).ToString();
            this.FailedTotal.Text = cumulativeResults.GetOutcomeCount(TestOutcome.Failed, false).ToString();
            this.BlockedTotal.Text = cumulativeResults.GetOutcomeCount(TestOutcome.Blocked, false).ToString();
            this.QueuedTotal.Text = cumulativeResults.GetOutcomeCount(TestOutcome.None, false).ToString();
            this.InconclusiveTotal.Text = cumulativeResults.GetOutcomeCount(TestOutcome.Inconclusive, false).ToString();
            this.NeverExecuted.Text = cumulativeResults.GetOutcomeCount(TestOutcome.Unspecified, false).ToString();
            this.ActiveTotal.Text = cumulativeResults.GetActiveCount(false).ToString();
            this.OtherTotal.Text = cumulativeResults.GetOtherCount(false).ToString();
            this.TotalLabel.Text = cumulativeResults.Count.ToString();
            this.PassedNonDupLabel.Text = cumulativeResults.GetOutcomeCount(TestOutcome.Passed).ToString();
            this.FailedNonDupLabel.Text = cumulativeResults.GetOutcomeCount(TestOutcome.Failed).ToString();
            this.BlockedNonDupLabel.Text = cumulativeResults.GetOutcomeCount(TestOutcome.Blocked).ToString();
            this.QueuedNonDupLabel.Text = cumulativeResults.GetOutcomeCount(TestOutcome.None).ToString();
            this.InconclusiveNonDupLabel.Text = cumulativeResults.GetOutcomeCount(TestOutcome.Inconclusive).ToString();
            this.NeverExecutedNonDupLabel.Text = cumulativeResults.GetOutcomeCount(TestOutcome.Unspecified).ToString();
            this.ActiveNonDupLabel.Text = cumulativeResults.GetActiveCount(true).ToString();
            this.OtherNonDupLabel.Text = cumulativeResults.GetOtherCount().ToString();
            this.TotalNonDupLabel.Text = cumulativeResults.NonDupCount.ToString();
        }

        /// <summary>
        /// Update the selected UI counts
        /// </summary>
        private void SelectedCounts(TestResult result)
        {
            List<TestResult> results = new List<TestResult>();
            results.Add(result);

            // Update the UI count
            UpdateSelectedCounts(new AllResults(results));
        }

        /// <summary>
        /// Update the selected UI counts
        /// </summary>
        private void SelectedCounts(TreeNode selectedNode)
        {
            List<TestResult> results = new List<TestResult>();

            // Get all the results and sub results
            this.GetResultList(results, selectedNode);

            // Update the UI count
            UpdateSelectedCounts(new AllResults(results));
        }

        /// <summary>
        /// Update the selected count
        /// </summary>
        /// <param name="cumulativeResults">The cumulative result count</param>
        private void UpdateSelectedCounts(AllResults cumulativeResults)
        {
            // Clear old values
            this.PassedSelectedLabel.Text = "0";
            this.FailedSelectedLabel.Text = "0";
            this.BlockedSelectedLabel.Text = "0";
            this.QueuedSelectedLabel.Text = "0";
            this.InconclusiveSelectedLabel.Text = "0";
            this.NotExecutedSelectedLabel.Text = "0";
            this.ActiveSelectedLabel.Text = "0";
            this.OtherSelectedLabel.Text = "0";
            this.TotalSelectedLabel.Text = "0";

            this.PassedSelectedLabel.Text = cumulativeResults.GetOutcomeCount(TestOutcome.Passed, false).ToString();
            this.FailedSelectedLabel.Text = cumulativeResults.GetOutcomeCount(TestOutcome.Failed, false).ToString();
            this.BlockedSelectedLabel.Text = cumulativeResults.GetOutcomeCount(TestOutcome.Blocked, false).ToString();
            this.QueuedSelectedLabel.Text = cumulativeResults.GetOutcomeCount(TestOutcome.None, false).ToString();
            this.InconclusiveSelectedLabel.Text = cumulativeResults.GetOutcomeCount(TestOutcome.Inconclusive, false).ToString();
            this.NotExecutedSelectedLabel.Text = cumulativeResults.GetOutcomeCount(TestOutcome.Unspecified, false).ToString();
            this.ActiveSelectedLabel.Text = cumulativeResults.GetActiveCount(false).ToString();
            this.OtherSelectedLabel.Text = cumulativeResults.GetOtherCount(false).ToString();
            this.TotalSelectedLabel.Text = cumulativeResults.Count.ToString();
        }

        /// <summary>
        /// Get a list of results
        /// </summary>
        /// <param name="results">List of test results</param>
        /// <param name="node">Node to add the results to</param>
        private void GetResultList(List<TestResult> results, TreeNode node)
        {
            // Loop over all the child nodes
            foreach (TreeNode singleNode in node.Nodes)
            {
                TestResult result = (TestResult)singleNode.Tag;
                if (result.TestID < 0)
                {
                    this.GetResultList(results, singleNode);
                }
                else
                {
                    results.Add(result);
                }
            }
        }

        /// <summary>
        /// Loop over the tree nodes and assure all the icons are set properly
        /// </summary>
        /// <param name="node">Node look over</param>
        private void SetIcons(TreeNode node)
        {
            TestResult result = (TestResult)node.Tag;
            if (result.TestID > 0)
                {
                    int index = SharedUtils.GetIconIndex(result.Outcome, result.State);
                    node.ImageIndex = index;
                    node.SelectedImageIndex = index;
                }
            else
            {
                foreach (TreeNode singleNode in node.Nodes)
                {
                    this.SetIcons(singleNode);
                }
            }
        }

        /// <summary>
        /// Save the current results
        /// </summary>
        /// <param name="filename">The name of the file to save the data to</param>
        private void SaveTree(string filename)
        {
            try
            {
                Stream file = File.Open(filename, FileMode.Create);
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(file, this.StructureView.Nodes.Cast<TreeNode>().ToList());
                file.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show("Failed to cache tree structure because:\r\n" + e.Message, "Caching Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Load a cached TreeView
        /// </summary>
        /// <param name="fileName">The file to load</param>
        /// <returns>True if the file was loaded</returns>
        private bool LoadTree(string fileName)
        {
            if (File.Exists(fileName))
            {
                FileInfo info = new FileInfo(fileName);
                string lastUpdate = info.LastWriteTime.ToString("f");

                try
                {
                    Stream file = File.Open(fileName, FileMode.Open);

                    try
                    {
                        BinaryFormatter binFormatter = new BinaryFormatter();
                        object obj = binFormatter.Deserialize(file);
                        this.StructureView.Nodes.AddRange((obj as IEnumerable<TreeNode>).ToArray());
                    }
                    catch
                    {
                        // Could not read the file so just act as though it does not exist
                        return false;
                    }
                    finally
                    {
                        file.Close();
                    }

                    // Get the last update time
                    this.ToolStripStatusLastUpdate.Text = lastUpdate;
                    this.StructureView.Tag = lastUpdate;

                    // Update the UI
                    this.UpdateTreeNodeIcons(this.StructureView.Nodes);
                    this.NewCounts();

                    return true;
                }
                catch (Exception e)
                {
                    MessageBox.Show("Failed to load cached tree structure because:\r\n" + e.Message, "Load Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            return false;
        }

        /// <summary>
        /// Update the icons for all nodes and subnodes in the TreeNodeCollection
        /// </summary>
        /// <param name="nodes">The TreeNodeCollection to update icons for</param>
        private void UpdateTreeNodeIcons(TreeNodeCollection nodes)
        {
            foreach (TreeNode node in nodes)
            {
                this.UpdateTreeNodeIcons(node);
            }
        }

        /// <summary>
        /// Update the tree icon for a specific tree node
        /// </summary>
        /// <param name="node">The node update, this includes any children of the node</param>
        private void UpdateTreeNodeIcons(TreeNode node)
        {
            TestResult result = (TestResult)node.Tag;

            // We only need to update the icon if we are dealing with a test
            if (result.TestID >= 0)
            {
                node.ImageIndex = SharedUtils.GetIconIndex(result.Outcome, result.State);
                node.SelectedImageIndex = node.ImageIndex;
            }

            // Update the children
            foreach (TreeNode subNode in node.Nodes)
            {
                this.UpdateTreeNodeIcons(subNode);
            }
        }

        /// <summary>
        /// Ad a single result to the results ListView
        /// </summary>
        /// <param name="result">The result to add</param>
        private void AddTestResult(TestResult result)
        {
            ListViewItem test = new ListViewItem(result.TestID.ToString());
            test.Text = SharedUtils.GetStateText(result.Outcome, result.State);
            test.ImageIndex = SharedUtils.GetIconIndex(result.Outcome, result.State);
            test.SubItems.Add(result.State.ToString());
            test.SubItems.Add(result.TestID.ToString());
            test.SubItems.Add(result.Name.ToString());
            test.SubItems.Add(result.Area.ToString());
            test.SubItems.Add(result.Priority.ToString());
            test.SubItems.Add(SharedUtils.FormattedDateTime(result.Completed));
            test.SubItems.Add(result.IsAutomated.ToString());
            test.Tag = result.SuiteID;

            if (result.TestCaseArea != null)
            {
                test.SubItems.Add(result.TestCaseArea.ToString());
            }
            else
            {
                test.SubItems.Add(string.Empty);
            }

            test.SubItems.Add(result.Configuration);

            if (!string.IsNullOrEmpty(result.FailureType))
            {
                test.SubItems.Add(result.FailureType);
            }
            else
            {
                test.SubItems.Add(string.Empty);
            }

            if (result.ErrorMessage != null)
            {
                test.SubItems.Add(result.ErrorMessage.ToString());
            }
            else
            {
                test.SubItems.Add(string.Empty);
            }


            test.SubItems.Add(SharedUtils.GetHistorical(result.HistoricOutcomes));

            this.ResultsListView.Items.Add(test);
        }

        /// <summary>
        /// Get unique file name for the given unformatted name
        /// </summary>
        /// <param name="unformattedName">The unformatted name, typically the plan name</param>
        /// <returns>A valid file name that matches the given formatted name</returns>
        private string NameSpecificFileName(string unformattedName)
        {
            return unformattedName.GetHashCode() + ".ttwtm";
        }

        /// <summary>
        /// Kill child threads
        /// </summary>
        private void KillQueryThreads()
        {
            foreach (Thread singleThread in this.childThreads)
            {
                if (singleThread.IsAlive)
                {
                    singleThread.Join(1000);
                    singleThread.Abort();
                }
            }

            // Clear out the thread list
            this.childThreads = new List<Thread>();
        }

        /// <summary>
        /// Open a specific test case in the web version of TFS
        /// </summary>
        /// <param name="testCaseId">The ID of the test case</param>
        /// <param name="suiteID">The suite the test was assoicated with</param>
        private void OpenWebTestCase(int testCaseId, int suiteID)
        {
            string url = connection.GetOpenSubpath(testCaseId, suiteID);
           // url = ConfigurationManager.AppSettings["WebRoot"] + connection.GetOpenSubpath(testCaseId, suiteID);

            SharedUtils.OpenExternalProcess(url, "Cannot open test");
            // StaticShared.OpenExternalProcess(string.Format("{0}/wi.aspx?id={1}", ConfigurationManager.AppSettings["WebRoot"], testCaseId), "Open Web TFS Failed");
        }
    }
}
