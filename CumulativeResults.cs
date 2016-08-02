using Microsoft.TeamFoundation.TestManagement.Client;
using System.Collections.Generic;

namespace TestPlanViewer
{
    /// <summary>
    /// Class to represent a collection of results
    /// </summary>
    public class AllResults
    {
        /// <summary>
        /// A dictionary of result lists
        /// </summary>
        private Dictionary<int, SortedSet<TestResult>> allResults;

        /// <summary>
        /// Initializes a new instance of the CumulativeResults class
        /// </summary>
        /// <param name="results">A list of test results</param>
        public AllResults(List<TestResult> results)
        {
            // Setup
            this.Count = 0;
            this.NonDupCount = 0;
            this.allResults = new Dictionary<int, SortedSet<TestResult>>();

            // Loop over every result
            foreach (TestResult result in results)
            {
                this.AddResults(result);
            }
        }

        /// <summary>
        /// Gets the total number of test results
        /// </summary>
        public int Count { get; private set; }

        /// <summary>
        /// Gets the number of test results without counting duplicates
        /// </summary>
        public int NonDupCount { get; private set; }

        /// <summary>
        /// Get a list of all the all the test results and if they are primary or duplicates
        /// </summary>
        /// <returns>The list oftest results and if they are primary or duplicates</returns>
        public List<KeyValuePair<bool, TestResult>> GetTypesAndResults()
        {
            // Make the list
            List<KeyValuePair<bool, TestResult>> resultsAndTypes = new List<KeyValuePair<bool, TestResult>>();

            // Loop over all the tests
            foreach (KeyValuePair<int, SortedSet<TestResult>> testResults in this.allResults)
            {
                bool isPrimary = true;
                foreach (TestResult singleResult in testResults.Value.Reverse())
                {
                    resultsAndTypes.Add(new KeyValuePair<bool, TestResult>(isPrimary, singleResult));
                    if (isPrimary)
                    {
                        isPrimary = !isPrimary;
                    }
                }
            }

            return resultsAndTypes;
        }

        /// <summary>
        /// Get the total number of results with uncommon outcomes
        /// </summary>
        /// <param name="latestOnly">Are we only looking at the latest - AKA no duplicates</param>
        /// <returns>The total number of results with uncommon outcomes</returns>
        public int GetOtherCount(bool latestOnly = true)
        {
            int total = 0;

            // Loop over all the tests
            foreach (KeyValuePair<int, SortedSet<TestResult>> testResults in this.allResults)
            {
                int lastIndex = testResults.Value.Count - 1;

                if (latestOnly)
                {
                    // Only look at the latest result
                    if (TestPointState.Ready != testResults.Value.Max.State && SharedUtils.IsOther(testResults.Value.Max.Outcome))
                    {
                        total++;
                    }
                }
                else
                {
                    foreach (TestResult singleResult in testResults.Value)
                    {
                        if (TestPointState.Ready != singleResult.State && SharedUtils.IsOther(singleResult.Outcome))
                        {
                            total++;
                        }
                    }
                }
            }

            return total;
        }

        /// <summary>
        /// Get the total number of active outcomes
        /// </summary>
        /// <param name="latestOnly">Are we only looking at the latest - AKA no duplicates</param>
        /// <returns>The total number of results with active outcomes</returns>
        public int GetActiveCount(bool latestOnly = true)
        {
            int total = 0;

            // Loop over all the tests
            foreach (KeyValuePair<int, SortedSet<TestResult>> testResults in this.allResults)
            {
                int lastIndex = testResults.Value.Count - 1;

                if (latestOnly)
                {
                    // Only look at the latest result
                    if (TestPointState.Ready == testResults.Value.Max.State)
                    {
                        total++;
                    }
                }
                else
                {
                    foreach (TestResult singleResult in testResults.Value)
                    {
                        if (TestPointState.Ready == singleResult.State)
                        {
                            total++;
                        }
                    }
                }
            }

            return total;
        }

        /// <summary>
        /// Get the total number of results with the given outcome
        /// </summary>
        /// <param name="outcomeType">The type of outcome we are looking foe</param>
        /// <param name="latestOnly">Are we only looking at the latest - AKA no duplicates</param>
        /// <returns>The total numbe of results with the given outcome</returns>
        public int GetOutcomeCount(TestOutcome outcomeType, bool latestOnly = true)
        {
            int total = 0;

            // Loop over all the tests
            foreach (KeyValuePair<int, SortedSet<TestResult>> testResults in this.allResults)
            {
                int lastIndex = testResults.Value.Count - 1;

                if (latestOnly)
                {
                    // Only look at the latest result
                    if (outcomeType == testResults.Value.Max.Outcome && TestPointState.Ready != testResults.Value.Max.State)
                    {
                        total++;
                    }
                }
                else
                {
                    foreach (TestResult singleResult in testResults.Value)
                    {
                        if (outcomeType == singleResult.Outcome && TestPointState.Ready != singleResult.State)
                        {
                            total++;
                        }
                    }
                }
            }

            return total;
        }

        /// <summary>
        /// Add a sigle result
        /// </summary>
        /// <param name="result">The single result to add</param>
        private void AddResults(TestResult result)
        {
            SortedSet<TestResult> currentResults;
            int testID = result.TestID;

            // Check if the test already exists
            if (this.allResults.ContainsKey(testID))
            {
                currentResults = this.allResults[testID];
            }
            else
            {
                // The test is new so add a new result list
                currentResults = new SortedSet<TestResult>(new ResultComparer());
                this.allResults.Add(result.TestID, currentResults);
                this.NonDupCount++;
            }

            currentResults.Add(result);
            this.Count++;
        }
    }
}
