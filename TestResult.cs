using Microsoft.TeamFoundation.TestManagement.Client;
using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using MSFailureType = Microsoft.TeamFoundation.TestManagement.Client.FailureType;

namespace TestPlanViewer
{
    /// <summary>
    /// Object to hold a single test result
    /// </summary>
    [Serializable()]
    public class TestResult : ISerializable
    {
        /// <summary>
        /// Initializes a new instance of the TestResult class for deserialization
        /// </summary>
        /// <param name="info">Serialization information</param>
        /// <param name="context">The streaming context</param>
        public TestResult(SerializationInfo info, StreamingContext context)
        {
            // Deserialization the info
            this.Outcome = (TestOutcome)info.GetValue("Outcome", typeof(TestOutcome));
            this.State = (TestPointState)info.GetValue("State", typeof(TestPointState));
            this.Name = (string)info.GetValue("Name", typeof(string));
            this.TestID = (int)info.GetValue("TestID", typeof(int));
            this.Completed = (DateTime)info.GetValue("Completed", typeof(DateTime));
            this.Duration = (TimeSpan)info.GetValue("Duration", typeof(TimeSpan));
            this.Configuration = (string)info.GetValue("Configuration", typeof(string));
            this.IsAutomated = (bool)info.GetValue("IsAutomated", typeof(bool));
            this.Priority = (int)info.GetValue("Priority", typeof(int));
            this.FailureType = (string)info.GetValue("FailureType", typeof(string));
            this.TestCaseArea = (string)info.GetValue("TestCaseArea", typeof(string));
            this.ErrorMessage = (string)info.GetValue("ErrorMessage", typeof(string));
            this.Exists = (bool)info.GetValue("Exists", typeof(bool));
            this.SuiteID = (int)info.GetValue("SuiteID", typeof(int));
            this.Area = (string)info.GetValue("Area", typeof(string));
            this.MostRecentRunID = (int)info.GetValue("MostRecentRunId", typeof(int));
            this.HistoricOutcomes = (List<TestOutcome>)info.GetValue("HistoricOutcomes", typeof(List<TestOutcome>));
        }

        /// <summary>
        /// Initializes a new instance of the TestResult class
        /// </summary>
        /// <param name="name">The test name</param>
        /// <param name="suiteID">The test's suite ID</param>
        /// <param name="testID">The test's ID</param>
        /// <param name="planArea">The test's plan area</param>
        /// <param name="config">The test's config</param>
        public TestResult(string name, int suiteID, int testID, string planArea, string config = "" )
        {
            this.Name = name;
            this.TestID = testID;
            this.SuiteID = suiteID;
            this.MostRecentRunID = -1;
            this.Area = planArea;
            this.Configuration = config;
            this.TestCaseArea = string.Empty;
            this.FailureType = string.Empty;
            this.ErrorMessage = string.Empty;
            this.HistoricOutcomes = new List<TestOutcome>();
        }

        /// <summary>
        /// Gets or sets the state of the test
        /// </summary>
        public TestPointState State { get; set; }

        /// <summary>
        /// Gets the outcome of the test
        /// </summary>
        public TestOutcome Outcome { get; private set; }

        /// <summary>
        /// Gets the name of the test
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Gets the path to the test's suite
        /// </summary>
        public string Area { get; private set; }

        /// <summary>
        /// Gets a value indicating whether the test exists
        /// </summary>
        public bool Exists { get; private set; }

        /// <summary>
        /// Gets a value indicating whether the test is automated
        /// </summary>
        public bool IsAutomated { get; private set; }

        /// <summary>
        /// Gets the test ID
        /// </summary>
        public int TestID { get; private set; }

        /// <summary>
        /// Gets the test suite ID
        /// </summary>
        public int SuiteID { get; private set; }

        /// <summary>
        /// Gets when the test was completed
        /// </summary>
        public DateTime Completed { get; private set; }

        /// <summary>
        /// Gets how long did the test take
        /// </summary>
        public TimeSpan Duration { get; private set; }

        /// <summary>
        /// Gets the TFS test area
        /// </summary>
        public string TestCaseArea { get; private set; }

        /// <summary>
        /// Gets the test's priority
        /// </summary>
        public int Priority { get; private set; }

        /// <summary>
        /// Gets the test's configuration
        /// </summary>
        public string Configuration { get; private set; }

        /// <summary>
        /// Gets the type of test failure
        /// </summary>
        public string FailureType { get; private set; }

        /// <summary>
        /// Gets the associated error message
        /// </summary>
        public string ErrorMessage { get; private set; }

        /// <summary>
        /// Gets the most recent run ID
        /// </summary>
        public int MostRecentRunID { get; private set; }

        /// <summary>
        /// List of historic results
        /// </summary>
        public List<TestOutcome> HistoricOutcomes { get; set; }

        /// <summary>
        /// Update the result for the test case result
        /// </summary>
        /// <param name="result">Test case results</param>
        /// <param name="state">The current state of the test</param>
        /// <param name="outcomes">List of last couple of outcomes</param>
        /// <param name="currentlyExists">Does the test case exist</param>
        public void UpdateResult(ITestCaseResult result, TestPointState state, List<TestOutcome> outcomes, bool currentlyExists)
        {
            ITestCase testcase = result.GetTestCase();

            this.MostRecentRunID = result.TestRunId;
            this.Outcome = result.Outcome;
            this.HistoricOutcomes = outcomes;
            this.Name = testcase.Title;
            this.TestID = testcase.Id;

            this.Configuration = result.TestConfigurationName;

            // Handle missing dates for test in progress
            this.Completed = result.Outcome == TestOutcome.None ? DateTime.Now : result.DateCompleted;
            this.State = state;
            this.Duration = result.Duration;
            this.IsAutomated = testcase.IsAutomated;
            UpdateBaseTestData(testcase);



            this.FailureType = result.FailureTypeId == (int)MSFailureType.NullValue ? string.Empty : ((MSFailureType)result.FailureTypeId).ToString();

            // FailureType
            this.ErrorMessage = result.ErrorMessage == null ? string.Empty : result.ErrorMessage.ToString();
            this.Exists = currentlyExists;
        }

        /// <summary>
        /// Update data related to the test item itself
        /// </summary>
        /// <param name="testCaseWorkItem">The test case work item</param>
        public void UpdateBaseTestData(ITestCase testCaseWorkItem)
        {
            this.TestCaseArea = testCaseWorkItem.Area;
            this.Priority = testCaseWorkItem.Priority;
        }

        /// <summary>
        /// Get a serialized version of the object
        /// </summary>
        /// <param name="info">Serialization information</param>
        /// <param name="context">The streaming context</param>
        public void GetObjectData(SerializationInfo info, StreamingContext context)
        {
            info.AddValue("HistoricOutcomes", this.HistoricOutcomes);
            info.AddValue("Outcome", this.Outcome);
            info.AddValue("State", this.State);
            info.AddValue("Name", this.Name);
            info.AddValue("SuiteID", this.SuiteID);
            info.AddValue("Configuration", this.Configuration);
            info.AddValue("TestID", this.TestID); 
            info.AddValue("MostRecentRunId", this.MostRecentRunID);
            info.AddValue("Completed", this.Completed);
            info.AddValue("Duration", this.Duration);
            info.AddValue("IsAutomated", this.IsAutomated);
            info.AddValue("Priority", this.Priority);
            info.AddValue("FailureType", this.FailureType);
            info.AddValue("TestCaseArea", this.TestCaseArea);
            info.AddValue("ErrorMessage", this.ErrorMessage);
            info.AddValue("Exists", this.Exists);
            info.AddValue("Area", this.Area);
        }
    }
}
