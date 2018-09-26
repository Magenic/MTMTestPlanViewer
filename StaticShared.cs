using Microsoft.TeamFoundation.TestManagement.Client;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace TestPlanViewer
{
    /// <summary>
    /// Class to hold shared util funtions
    /// </summary>
    public static class SharedUtils
    {        
        /// <summary>
        /// Applications common time span fomat
        /// </summary>
        public const string TimeSapnFormat = "g";

        /// <summary>
        /// Applications common date fomat
        /// </summary>
        public const string DateTimeFormat = "yyyy/MM/dd";

        /// <summary>
        /// Get the icon index given the test outcome and state
        /// </summary>
        /// <param name="outcome">The test outcome</param>
        /// <param name="state">The state of the test</param>
        /// <returns>The icon index for the outcome/state</returns>
        public static int GetIconIndex(TestOutcome outcome, TestPointState state)
        {
            if (state == TestPointState.Ready)
            {
                return 2;
            }

            switch (outcome)
            {
                case TestOutcome.Passed:
                    return 3;
                case TestOutcome.Failed:
                    return 4;
                case TestOutcome.NotExecuted:
                case TestOutcome.Inconclusive:
                case TestOutcome.Timeout:
                case TestOutcome.Warning:
                    return 5;
                case TestOutcome.Blocked:
                    return 6;
                case TestOutcome.None:
                    return 7;
                case TestOutcome.Unspecified:
                    return 8;
                default:
                    return 1;
            }
        }

        /// <summary>
        /// Get the historical string value
        /// </summary>
        /// <param name="outcomes">List of outcomes - most recent to least</param>
        /// <returns>A historical string</returns>
        public static string GetHistorical(List<TestOutcome> outcomes)
        {
            // Return the empty string if the outcome list is null or empty
            if (outcomes == null || outcomes.Count == 0)
            {
                return string.Empty;
            }

            string results = string.Empty;
            foreach (TestOutcome outcome in outcomes)
            {
                results += outcome.ToString() + " ";
            }

            return results;
        }

        /// <summary>
        /// Get the proper outcome text given the test outcome and state
        /// </summary>
        /// <param name="outcome">The test outcome</param>
        /// <param name="state">The state of the test</param>
        /// <returns>The icon index that corresponds to the outcome</returns>
        public static string GetStateText(TestOutcome outcome, TestPointState state)
        {
            if (state == TestPointState.Ready)
            {
                return "Active";
            }

            return outcome == TestOutcome.Unspecified ? "Unspecified (Never Executed)" : outcome.ToString();
        }

        /// <summary>
        /// Format a date time
        /// </summary>
        /// <param name="dateTime">The date time to format</param>
        /// <returns>A formatted DateTime string</returns>
        public static string FormattedDateTime(DateTime dateTime)
        {
            return dateTime == DateTime.MinValue ? "Not run" : dateTime.ToString(DateTimeFormat);
        }

        /// <summary>
        /// Format a time span
        /// </summary>
        /// <param name="timeSpan">The time span to format</param>
        /// <returns>A formatted TimeSpan string</returns>
        public static string FormattedTimeSpan(TimeSpan timeSpan)
        {
            return timeSpan == TimeSpan.Zero ? "NA" : timeSpan.ToString(TimeSapnFormat);
        }

        /// <summary>
        /// Check if this outcome is one of the outcomes we don't explicitly deal with
        /// </summary>
        /// <param name="testOutcome">The test outcome</param>
        /// <returns>Is this one of the other test results</returns>
        public static bool IsOther(TestOutcome testOutcome)
        {
            // Check if the outcome is one of the common types
            switch (testOutcome)
            {
                case TestOutcome.Passed:
                case TestOutcome.Failed:
                case TestOutcome.Blocked:
                case TestOutcome.Inconclusive:
                case TestOutcome.Unspecified:
                case TestOutcome.None:
                    return false;
                default:
                    return true;
            }
        }
 
        /// <summary>
        /// Get a file from the user
        /// </summary>
        /// <param name="filter">Save filed dialog filter</param>
        /// <param name="defaultExtention">The default file extension</param>
        /// <returns>The file the user wants to create or an empty string if they didn't select a file</returns>
        public static string GetFileName(string filter, string defaultExtention)
        {
            try
            {
                // Launch save file dialog and get file name
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = filter;
                saveFileDialog.DefaultExt = defaultExtention;
                saveFileDialog.FilterIndex = 2;
                saveFileDialog.RestoreDirectory = true;

                // If the user selected a file use it
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return saveFileDialog.FileName;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Failed to save file because:\r\n" + e.Message, "Save Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // No file was selected so return the empty string
            return string.Empty;
        }

        /// <summary>
        /// Launch an external process
        /// </summary>
        /// <param name="processName">The process name. This can be an exe or file with a known extension</param>
        /// <param name="errorCaption">If there is a failure what should the error caption read</param>
        public static void OpenExternalProcess(string processName, string errorCaption)
        {
            try
            {
                Process.Start(processName);
            }
            catch (Exception exception)
            {
                MessageBox.Show(string.Format("Failed launch '{0}' because:\r\n{1}", processName, exception.Message), errorCaption, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Save the results to a tab delimited file
        /// </summary>
        /// <param name="results">A collection of our test results</param>
        /// <param name="fileName">The name of the file to create</param>
        public static void CreateResultFlatFile(AllResults results, string fileName)
        {
            try
            {
                // Save the file
                StreamWriter outputStream;
                if ((outputStream = new StreamWriter(fileName)) != null)
                {
                    // Setup header row
                    outputStream.WriteLine("Outcome\tName\tArea\tIsAutomated\tCurrentlyExists\tID\tPriority\tCompleted\tDuration\tTestCaseArea\tFailureType\tErrorMessage\tPrimary\tState\tRecent History (*most recent first)");

                    foreach (KeyValuePair<bool, TestResult> result in results.GetTypesAndResults())
                    {
                        // Get the specific result
                        TestResult singleResult = result.Value;

                        StringBuilder singleLine = new StringBuilder();
                        singleLine.Append(SharedUtils.GetStateText(singleResult.Outcome, singleResult.State));
                        singleLine.Append("\t");
                        singleLine.Append(singleResult.Name.Replace("\t", " "));
                        singleLine.Append("\t");
                        singleLine.Append(singleResult.Area.Replace("\t", " "));
                        singleLine.Append("\t");
                        singleLine.Append(singleResult.IsAutomated);
                        singleLine.Append("\t");
                        singleLine.Append(singleResult.Exists);
                        singleLine.Append("\t");
                        singleLine.Append(singleResult.TestID);
                        singleLine.Append("\t");
                        singleLine.Append(singleResult.Priority);
                        singleLine.Append("\t");
                        singleLine.Append(SharedUtils.FormattedDateTime(singleResult.Completed));
                        singleLine.Append("\t");
                        singleLine.Append(SharedUtils.FormattedTimeSpan(singleResult.Duration));
                        singleLine.Append("\t");
                        singleLine.Append(singleResult.TestCaseArea.Replace("\t", " "));
                        singleLine.Append("\t");
                        singleLine.Append(singleResult.FailureType.Replace("\t", " "));
                        singleLine.Append("\t");
                        singleLine.Append(singleResult.ErrorMessage.Replace("\t", " ").Replace("\r\n", "  "));
                        singleLine.Append("\t");
                        singleLine.Append(result.Key.ToString());
                        singleLine.Append("\t");
                        singleLine.Append(singleResult.State);
                        singleLine.Append("\t");
                        singleLine.Append(SharedUtils.GetHistorical(singleResult.HistoricOutcomes));
                        outputStream.WriteLine(singleLine.ToString());
                    }

                    // Close the file
                    outputStream.Close();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("Failed to save file because:\r\n" + e.Message, "Save Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
