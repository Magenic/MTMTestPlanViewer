using System.Collections.Generic;

namespace TestPlanViewer
{
    /// <summary>
    /// ResultComparer object
    /// </summary>
    public class ResultComparer : Comparer<TestResult>
    {
        /// <summary>
        /// Compare test results
        /// </summary>
        /// <param name="testOne">The first item</param>
        /// <param name="testTwo">The second item</param>
        /// <returns>The compare int value</returns>
        public override int Compare(TestResult testOne, TestResult testTwo)
        {
            if (testOne.Completed == testTwo.Completed)
            {
                return testOne.SuiteID.CompareTo(testTwo.SuiteID);
            }
            else
            {
                return testOne.Completed.CompareTo(testTwo.Completed);
            }
        }
    }
}
