using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
namespace TestPlanViewer
{
    /// <summary>
    /// Object to handle communication with the TFS results DB
    /// </summary>
    public class TFSConnection : IDisposable
    {
        /// <summary>
        /// The current test plan we are connected to
        /// </summary>
        private ITestPlan currentTestPlan;

        /// <summary>
        /// The current project collection
        /// </summary>
        private TfsTeamProjectCollection tfsProjCollection;

        private bool _Disposed;

        /// <summary>
        /// Cleanup
        /// </summary>
        /// <param name="disposing">Should displose</param>
        private void Dispose(bool disposing)
        {
            if (_Disposed)
            {
                return;
            }
            if (disposing)
            {
                tfsProjCollection.Dispose();
            }

            _Disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Initializes a new instance of the TFSConnection class
        /// </summary>
        public TFSConnection()
        {
            this.View = new TreeView();
            this.SetupConnection();
        }

        /// <summary>
        /// Get the sub path to the given test
        /// </summary>
        /// <param name="testCaseID">The test case ID</param>
        /// <param name="testSuiteID">The suite the test case is in</param>
        /// <returns>The subpath - if we can return the test results, if not return the workitem</returns>
        public string GetOpenSubpath(int testCaseID, int testSuiteID)
        {
            // Get the root url
            string path = tfsProjCollection.Uri.ToString().Replace("http:", "mtm:").Replace("https:", "mtm:");

            StringBuilder openPath = new StringBuilder();
            openPath.Append("/p:");
            openPath.Append(currentTestPlan.AreaPath);
            openPath.Append("/Testing/");

            ITestPointCollection points = currentTestPlan.QueryTestPoints("SELECT * FROM TestPoint WHERE TestCaseID='" + testCaseID + "' and SuiteId='" + testSuiteID + "'");

            // If there are no results choose the work item
            if (points.Count == 0 || points[0].MostRecentResult == null)
            {
                openPath.Append("workitem/open?id=" + testCaseID);
            }
            else
            {
                // Open the results for the test
                openPath.Append("testresult/open?id=");
                openPath.Append(points[0].MostRecentResult.TestResultId);
                openPath.Append("&runid=");
                openPath.Append(points[0].MostRecentResult.TestRunId);
            }

            return Uri.EscapeUriString(path + openPath.ToString());
        }

        /// <summary>
        /// Initializes a new instance of the TFSConnection class
        /// </summary>
        /// <param name="view">The tree view object to update</param>
        public TFSConnection(TreeView view)
        {
            this.View = view;
            this.SetupConnection();
        }

        /// <summary>
        /// Gets or sets the active TreeView
        /// </summary>
        public TreeView View { private get; set; }

        /// <summary>
        /// Has the user successfully authenticated, if not try
        /// </summary>
        /// <returns>True if the user is authenticated</returns>
        public bool AuthenticationCheck()
        {
            try
            {
                // We are already connected
                if (this.tfsProjCollection.HasAuthenticated)
                {
                    return true;
                }
                else
                {
                    // Not authenticated so fix our connection
                    this.SetupConnection();
                }
            }
            catch
            {
                // Couldn't even check our connection so fix the connection
                this.SetupConnection();
            }

            // Are we authenticated
            return this.tfsProjCollection.HasAuthenticated;
        }

        /// <summary>
        /// Setup the connection to TFS DB results
        /// </summary>
        /// <param name="useStudioCon">Use the existing studio connection info</param>
        public void SetupConnection()
        {
            // Connect to server
            try
            {
                TeamProjectPicker teamProjectPicker = new TeamProjectPicker(TeamProjectPickerMode.SingleProject, false);
                DialogResult dlgResult = teamProjectPicker.ShowDialog();

                if (teamProjectPicker.SelectedTeamProjectCollection != null && dlgResult == DialogResult.OK)
                {
                    tfsProjCollection = teamProjectPicker.SelectedTeamProjectCollection;
                    ITestManagementService testManagementService = (ITestManagementService)tfsProjCollection.GetService(typeof(ITestManagementService));
                    WorkItemStore store = (WorkItemStore)tfsProjCollection.GetService(typeof(WorkItemStore));

                    string name = teamProjectPicker.SelectedProjects[0].Name;
                    ITestManagementTeamProject testProject = testManagementService.GetTeamProject(name);
                    if (testProject != null)
                    {
                        // Todo - Make this better
                        SelectPlanDialog planSelect = new SelectPlanDialog(GetPlans(testProject));
                        planSelect.StartPosition = FormStartPosition.CenterParent;

                        if (planSelect.ShowDialog() != DialogResult.OK)
                        {
                            return;
                        }

                        foreach (ITestPlan plan in testProject.TestPlans.Query("Select * From TestPlan"))
                        {
                            if (plan.Name.Equals(planSelect.Plan))
                            {
                                this.SetTestPlan(plan);
                                break;
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Error: The Test Project was not correct");
                    }
                }
                else
                {
                    return;
                }

                if (!this.EnsureAuthenticated())
                {
                    MessageBox.Show("Failed to authenticate to server.", "Authentication Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            catch (Exception err)
            {
                MessageBox.Show(string.Format("Failed to connect because:\r\n{0}", err.Message), "Server Connection Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }

        /// <summary>
        /// Get a list of plans
        /// </summary>
        /// <param name="testProject">The plan project</param>
        /// <returns>A list of test plans</returns>
        private List<string> GetPlans(ITestManagementTeamProject testProject)
        {
            List<string> planNames = new List<string>();
            this.EnsureAuthenticated();

            foreach (ITestPlan plan in testProject.TestPlans.Query("Select * From TestPlan"))
            {
                planNames.Add(plan.Name);
            }

            return planNames;
        }

        /// <summary>
        /// Set the test plan
        /// </summary>
        /// <param name="plan">The test plan</param>
        public void SetTestPlan(ITestPlan plan)
        {
            this.currentTestPlan = plan;
        }

        /// <summary>
        /// Get the test plan name
        /// </summary>
        public string GetTestPlanName()
        {
            if (this.currentTestPlan != null)
            {
                return this.currentTestPlan.Name;
            }

            return "";
        }

        /// <summary>
        /// Get the test suite structure
        /// </summary>
        /// <param name="includeLatestResult">Should we include test results</param>
        public void GetSuiteStructure(bool includeLatestResult = false)
        {
            // Get the existing values if we need them
            Dictionary<string, TestResult> exitingResults = new Dictionary<string, TestResult>();
            if (includeLatestResult)
            {
                foreach (TreeNode node in this.View.Nodes)
                {
                    this.GetNodeDictionary(node, exitingResults);
                }
            }

            // Get empty lists
            Dictionary<string, TreeNode> caseDicationary = new Dictionary<string, TreeNode>();
            this.View.Nodes.Clear();

            // Loop through the test suite to read the associated test case
            TreeNode newNode = new TreeNode(this.currentTestPlan.Name, 0, 0, this.GetTestNodes(this.currentTestPlan.RootSuite, this.currentTestPlan.RootSuite.Title, caseDicationary).ToArray());
            this.View.Nodes.Add(newNode);

            // Make sure we are not using a cached suite
            this.currentTestPlan.RootSuite.Refresh();

            // Loop over all the suites under the given suite
            foreach (ITestSuiteBase suite in this.currentTestPlan.RootSuite.SubSuites)
            {
                this.RecurseOnSuiteStructure(suite, suite.Title, newNode, caseDicationary);
            }

            // Ge the test resuls
            if (includeLatestResult)
            {
                this.UpdateAllTests(caseDicationary, exitingResults);
            }

            // Add suite result data
            newNode.Tag = new TestResult(string.Empty, this.currentTestPlan.RootSuite.Id, -1, this.currentTestPlan.RootSuite.Title);

            // Make sure it is sorted
            this.View.Sort();
        }

        /// <summary>
        /// Update a tree node with the latest test results
        /// </summary>
        /// <param name="viewNode">The suite node</param>
        public void UpdateWithLatestResult(TreeNode viewNode)
        {
            // don't try to update a null node
            if (viewNode == null)
            {
                return;
            }

            TestResult result = (TestResult)viewNode.Tag;

            // This is a Suite
            if (result.TestID < 0)
            {
                // Update all the test cases in the suite
                foreach (ITestPoint point in this.currentTestPlan.QueryTestPoints(string.Format("SELECT * FROM TestPoint WHERE SuiteId = {0} ", result.SuiteID)))
                {
                    ITestCase testCase = point.TestCaseWorkItem;

                    TreeNode updateNode = this.FindChildNode(viewNode.Nodes, point.TestCaseId, point.ConfigurationName);

                    if (updateNode != null)
                    {
                        // We already have the latest result
                        if (point.MostRecentRunId == ((TestResult)updateNode.Tag).MostRecentRunID || point.MostRecentRunId == 0)
                        {
                            ///Do somehting here
                            continue;
                        }
                    }

                    ITestCaseResult latestResult = point.MostRecentResult;


                    if (latestResult != null)
                    {
                        this.UpdateSingleNode(latestResult, testCase, point.State, updateNode);
                    }
                }
            }
            else
            {
                bool foundPoint = false;

                // Updating a sigle test case
                foreach (ITestPoint point in this.currentTestPlan.QueryTestPoints(string.Format("SELECT * FROM TestPoint WHERE SuiteId = {0} AND TestCaseId = {1} ", result.SuiteID, result.TestID)))
                {
                    if (point.ConfigurationName.Equals(result.Configuration))
                    {
                        foundPoint = true;
                        ITestCaseResult latestResult = point.MostRecentResult;

                        if (latestResult != null)
                        {
                            
                            this.UpdateSingleNode(latestResult, point.TestCaseWorkItem, point.State, viewNode);
                            break;
                        }
                    }
                }

                // The test was removed
                if (!foundPoint)
                {
                    TreeNode parent = viewNode.Parent;
                    viewNode.Remove();

                    if (parent.Nodes.Count == 0)
                    {
                        parent.Remove();
                    }
                }
            }
        }

        /// <summary>
        /// If the user is not connected try to connect
        /// </summary>
        /// <returns>Is the user authenticated</returns>
        private bool EnsureAuthenticated()
        {
            try
            {
                this.tfsProjCollection.EnsureAuthenticated();
                return this.tfsProjCollection.HasAuthenticated;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Recurse down the suite structure to setup a representative TreeView
        /// </summary>
        /// <param name="suite">The parent suite</param>
        /// <param name="parentPath">What is the suite path to the parent suite</param>
        /// <param name="viewNode">The current tree node we are adding to</param>
        /// <param name="caseDicationary">A dictionary to hold tree nodes for each suite/test combo</param>
        private void RecurseOnSuiteStructure(ITestSuiteBase suite, string parentPath, TreeNode viewNode, Dictionary<string, TreeNode> caseDicationary)
        {
            // Make sure we are not using a cached suite
            suite.Refresh();

            // Add a new node for the suite
            TreeNode newNode = new TreeNode(suite.Title, 0, 0, this.GetTestNodes(suite, parentPath, caseDicationary).ToArray());
            viewNode.Nodes.Add(newNode);
            newNode.Tag = new TestResult(string.Empty, suite.Id, -1, parentPath);

            // Add subsuites
            if (suite is IStaticTestSuite)
            {
                foreach (ITestSuiteBase childSuite in ((IStaticTestSuite)suite).SubSuites)
                {
                    this.RecurseOnSuiteStructure(childSuite, string.Format(@"{0}\{1}", parentPath, childSuite.Title), newNode, caseDicationary);

                }
            }
        }

        /// <summary>
        /// Get a list of test result tree nodes
        /// </summary>
        /// <param name="suite">The parent suite</param>
        /// <param name="parentPath">What is the suite path to the parent suite</param>
        /// <param name="caseDicationary">A dictionary to hold tree nodes for each suite/test combo</param>
        /// <returns>The list of test result tree nodes</returns>
        private List<TreeNode> GetTestNodes(ITestSuiteBase suite, string parentPath, Dictionary<string, TreeNode> caseDicationary)
        {
            // Make sure we are not using a cached suite
            suite.Refresh();

            List<TreeNode> testCaseNodes = new List<TreeNode>();

            // Loop over all the tast cases in the suite
            foreach (ITestSuiteEntry test in suite.TestCases)
            {
                FieldCollection fields = test.TestCase.CustomFields;

                foreach (IdAndName config in test.Configurations)
                {
                    // Setup a tree node
                    TreeNode node = new TreeNode(test.Title, 1, 1);
                    node.Tag = new TestResult(test.Title, suite.Id, test.Id, parentPath, config.Name);

                    // Add to our dicitonary of suites/test cases
                    caseDicationary.Add(this.TestCaseResultKey(suite.Id, test.Id, config.Name), node);

                    // Add a child tree node
                    testCaseNodes.Add(node);
                }
            }

            return testCaseNodes;
        }

        /// <summary>
        /// Recusively get child nodes and add the associated TestResults to result to a dictionary
        /// </summary>
        /// <param name="inspectNode">Node to inspect</param>
        /// <param name="resultsDicationary">Dictionary of results for specific tests in specific suites</param>
        private void GetNodeDictionary(TreeNode inspectNode, Dictionary<string, TestResult> resultsDicationary)
        {
            foreach (TreeNode childNode in inspectNode.Nodes)
            {
                TestResult result = (TestResult)childNode.Tag;
                if (result == null || result.TestID < 0)
                {
                    // This is a suite so recuse on it
                    this.GetNodeDictionary(childNode, resultsDicationary);
                }
                else
                {
                    // This is a test so 
                    resultsDicationary.Add(this.TestCaseResultKey(result.SuiteID, result.TestID, result.Configuration), result);
                }
            }
        }

        /// <summary>
        /// Update test cases with results
        /// </summary>
        /// <param name="caseDicationary">A dictionary to hold tree nodes for each suite/test combo</param>
        /// <param name="oldResultsDicationary">Dictionary of previous test results</param>
        private void UpdateAllTests(Dictionary<string, TreeNode> caseDicationary, Dictionary<string, TestResult> oldResultsDicationary)
        {
            // Query the server for the latest test results
            foreach (ITestPoint point in this.currentTestPlan.QueryTestPoints("SELECT * FROM TestPoint WHERE PlanId = " + this.currentTestPlan.Id))
            {
                if (!point.TestCaseExists)
                {
                    continue;
                }

                List<TestOutcome> outcomes = this.HistoricOutcomes(point.Id);

                string key = this.TestCaseResultKey(point.SuiteId, point.TestCaseId, point.ConfigurationName);
                if (oldResultsDicationary.ContainsKey(key) && oldResultsDicationary[key].Outcome != TestOutcome.None && point.MostRecentRunId == oldResultsDicationary[key].MostRecentRunID)
                {
                    this.UpdateSingleNode(oldResultsDicationary[key], point.State, point.TestCaseWorkItem, caseDicationary[key]);
                }
                else
                {
                    ITestCaseResult latestResult = point.MostRecentResult;

                    if (latestResult != null && caseDicationary.ContainsKey(key))
                    {
                        this.UpdateSingleNode(latestResult, point.TestCaseWorkItem, point.State, caseDicationary[key]);
                    }
                    else if (latestResult == null && oldResultsDicationary.ContainsKey(key) && outcomes.Count != ((TestResult)oldResultsDicationary[key]).HistoricOutcomes.Count)
                    {
                        ((TestResult)oldResultsDicationary[key]).HistoricOutcomes = this.HistoricOutcomes(point.Id);
                        this.UpdateSingleNode(oldResultsDicationary[key], point.State, point.TestCaseWorkItem, caseDicationary[key]);
                    }
                }

                if (caseDicationary.ContainsKey(key))
                {
                    UpdateSingleNodeWithBaseInfo(point.TestCaseWorkItem, caseDicationary[key]);
                }

            }
        }

        /// <summary>
        /// Update the test results for a single tree node
        /// </summary>
        /// <param name="singleResult">The test case result to update the node with</param>
        /// <param name="state">The current state of the test</param>
        /// <param name="viewNode">The tree node to update</param>
        private void UpdateSingleNode(ITestCaseResult singleResult, ITestCase testCase, TestPointState state, TreeNode viewNode)
        {
            TestResult result = (TestResult)viewNode.Tag;

            // No result so just return
            if (singleResult == null)
            {
                return;
            }
            else
            {
                List<TestOutcome> outcomes = this.HistoricOutcomes(singleResult.TestPointId);

                // Update the underlying test result
                result.UpdateResult(singleResult, state, outcomes, true);

                // Set the correct icon given the test outcome
                viewNode.ImageIndex = SharedUtils.GetIconIndex(singleResult.Outcome, state);
                viewNode.SelectedImageIndex = viewNode.ImageIndex;
            }
        }

        /// <summary>
        /// Update the test results for a single tree node
        /// </summary>
        /// <param name="singleResult">The test case result to update the node with</param>
        /// <param name="state">The current state of the test</param>
        /// <param name="viewNode">The tree node to update</param>
        private void UpdateSingleNodeWithBaseInfo(ITestCase testCase, TreeNode viewNode)
        {
            TestResult result = (TestResult)viewNode.Tag;
            result.UpdateBaseTestData(testCase);
        }

        /// <summary>
        /// Update the test results for a single tree node
        /// </summary>
        /// <param name="singleResult">The test case result to update the node with</param>
        /// <param name="state">The current state of the test</param>
        /// <param name="viewNode">The tree node to update</param>
        private void UpdateSingleNode(TestResult singleResult, TestPointState state, ITestCase testCase, TreeNode viewNode)
        {
            singleResult.State = state;

            // Update the underlying test result
            viewNode.Tag = singleResult;

            // Set the icon index for the UI
            viewNode.ImageIndex = SharedUtils.GetIconIndex(singleResult.Outcome, singleResult.State);
            viewNode.SelectedImageIndex = viewNode.ImageIndex;
        }

        /// <summary>
        /// Get the tree node that corresponds to the given test case result
        /// </summary>
        /// <param name="collection">The collection of tree nodes</param>
        /// <param name="latestResult">The test case result</param>
        /// <returns>The matching tree node</returns>
        private TreeNode FindChildNode(TreeNodeCollection collection, ITestCaseResult latestResult)
        {
            return this.FindChildNode(collection, latestResult.TestCaseId, latestResult.TestConfigurationName);
        }

        /// <summary>
        /// Get the tree node that corresponds to the given test case ID
        /// </summary>
        /// <param name="collection">The collection of tree nodes</param>
        /// <param name="testCaseId">The test case ID</param>
        /// <returns>The matching tree node</returns>
        private TreeNode FindChildNode(TreeNodeCollection collection, int testCaseId, string config)
        {
            foreach (TreeNode node in collection)
            {
                TestResult testNode = ((TestResult)node.Tag);

                if (testNode.TestID == testCaseId && testNode.Configuration.Equals(config))
                {
                    return node;
                }
            }

            return null;
        }

        /// <summary>
        /// Get the list of historical outcomes for a test point
        /// </summary>
        /// <param name="pointId">Test point to get outcomes for</param>
        /// <returns>List of outcomes, most recent to least</returns>
        private List<TestOutcome> HistoricOutcomes(int pointId)
        {
            List<TestOutcome> outcomes = new List<TestOutcome>();
            int loops = 0;

            // List of results starting with the most recent
            ITestCaseResultCollection testResults = this.currentTestPlan.Project.TestResults.Query("select * from TestResult where TestPointId = " + pointId + " order by DateStarted desc ");

            // Loop over test results
            foreach (ITestCaseResult result in testResults)
            {
                loops++;
                outcomes.Add(result.Outcome);

                if (loops >= 5)
                {
                    break;
                }
            }

            return outcomes;
        }

        /// <summary>
        /// Create a suite and test case specific key
        /// </summary>
        /// <typeparam name="T">The ID types</typeparam>
        /// <param name="suiteID">The suite ID</param>
        /// <param name="testCaseID">The test case ID</param>
        /// <param name="configID">The test configuration</param>
        /// <returns>A suite-test case specific key</returns>
        private string TestCaseResultKey<T>(T suiteID, T testCaseID, string config)
        {
            return string.Format("{0}-{1}-{2}", suiteID, testCaseID, config);
        }
    }
}