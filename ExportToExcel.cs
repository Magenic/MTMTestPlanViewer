using Microsoft.Office.Interop.Excel;
using Microsoft.TeamFoundation.TestManagement.Client;
using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;

namespace TestPlanViewer
{
    /// <summary>
    /// Class for creating excel documents
    /// </summary>
    public class ExportToExcel
    {
        /// <summary>
        /// A handle to the excel application
        /// </summary>
        private ExcelApp app = null;

        /// <summary>
        /// A handle to the excel application's worksheet
        /// </summary>
        private Worksheet worksheet = null;

        /// <summary>
        /// Launch a speadsheet with the cumulative content
        /// </summary>
        /// <param name="results">The cumulative results</param>
        public void LaunchExcel(AllResults results)
        {
            // Get the file name
            string file = SharedUtils.GetFileName("Excel Open XML | *.xlsm", "xlsm");

            if (string.IsNullOrEmpty(file))
            {
                // No file was provided so quite
                return;
            }

            this.LaunchExcel(results, file);
        }

        /// <summary>
        /// Launch a speadsheet with the cumulative content
        /// </summary>
        /// <param name="results">The cumulative content</param>
        /// <param name="fileDestination">File to save the content to</param>
        public void LaunchExcel(AllResults results, string fileDestination)
        {
            // Make sure the destination file has the correct extension
            if (!fileDestination.EndsWith(".xlsm", StringComparison.CurrentCultureIgnoreCase))
            {
                fileDestination = fileDestination + ".xlsm";
            }

            // Remove the existing file to avoid mystery 
            if (File.Exists(fileDestination))
            {
                try
                {
                    File.SetAttributes(fileDestination, FileAttributes.Normal);
                    File.Delete(fileDestination);
                }
                catch (Exception e)
                {
                    MessageBox.Show("Failed create file '" + fileDestination + "' because :\r\n" + e.Message, "Create Excel Document Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }

            // Create the instance of excel we will be interacting with
            if (!this.CreateHiddenExcelInstance())
            {
                return;
            }

            // Load the data into a temp file
            string tempFile = Path.GetTempFileName();
            SharedUtils.CreateResultFlatFile(results, tempFile);

            // Load the temp file
            this.app.Workbooks.OpenText(tempFile, Type.Missing, 1, XlTextParsingType.xlDelimited, XlTextQualifier.xlTextQualifierNone, Type.Missing, true, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // Format the active sheet
            this.FormatSheet(results);

            // Silently save the file
            this.app.ActiveWorkbook.SaveAs(fileDestination, XlFileFormat.xlOpenXMLWorkbookMacroEnabled, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // Close the copy of excel that we have an programmatic reference to
            this.CloseHiddenExcelInstance();

            // Cleanup the temp file
            File.Delete(tempFile);

            // Open excel in an external process
            SharedUtils.OpenExternalProcess(fileDestination, "Failed To Open Excel File");
        }

        /// <summary>
        /// Create a hidden instance of excel
        /// </summary>
        /// <returns>True if an instance of excel was created</returns>
        private bool CreateHiddenExcelInstance()
        {
            try
            {
                // Create a programmatic instance of excel and make sure it is not visible
                this.app = new ExcelApp();
                this.app.Visible = false;
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show("Failed create an excel document because:\r\n" + e.Message, "Create Excel Document Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return false;
        }

        /// <summary>
        /// Close the hidden instance of excel
        /// </summary>
        private void CloseHiddenExcelInstance()
        {
            // Close the copy of excel that we have an programmatic reference to
            if (this.app != null)
            {
                try
                {
                    Workbook workbook = this.app.ActiveWorkbook;
                    this.app.Workbooks.Close();
                    this.app.Quit();
                    Marshal.ReleaseComObject(this.worksheet);
                    Marshal.ReleaseComObject(workbook);
                    Marshal.ReleaseComObject(this.app);
                    Marshal.FinalReleaseComObject(this.app);
                }
                catch (Exception e)
                {
                    MessageBox.Show("Failed close hidden copy of excel because:\r\n" + e.Message, "Close Excel Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /// <summary>
        /// Format the spreadsheet
        /// </summary>
        /// <param name="results">The cumulative results</param>
        private void FormatSheet(AllResults results)
        {
            this.worksheet = this.app.ActiveSheet;
            this.SetTotals(results);
            this.FormattingDuplicates(results.Count);
            this.FormatHeaderRow();
        }

        /// <summary>
        /// Set the result totals
        /// </summary>
        /// <param name="results">The cumulative results</param>
        private void SetTotals(AllResults results)
        {
            int topRow = 3;
            int leftStart = 17;
            int leftStartTotalResults = leftStart + 1;
            int leftStartLatestResults = leftStart + 2;

            // Creates the main header
            this.CreateDataCell(topRow, leftStart, "Counts", Color.LightBlue);
            this.CreateDataCell(topRow, leftStartTotalResults, "All", Color.LightYellow);
            this.CreateDataCell(topRow, leftStartLatestResults, "Latest", Color.LightYellow);
            this.CreateDataCell(topRow + 1, leftStart, "Totals", Color.LightGray);
            this.CreateDataCell(topRow + 2, leftStart, "Passed", Color.LightGray);
            this.CreateDataCell(topRow + 3, leftStart, "Failed", Color.LightGray);
            this.CreateDataCell(topRow + 4, leftStart, "Blocked", Color.LightGray);
            this.CreateDataCell(topRow + 5, leftStart, "Active", Color.LightGray);
            this.CreateDataCell(topRow + 6, leftStart, "NotExecuted", Color.LightGray);
            this.CreateDataCell(topRow + 7, leftStart, "Inconclusive", Color.LightGray);
            this.CreateDataCell(topRow + 8, leftStart, "Unspecified", Color.LightGray);
            this.CreateDataCell(topRow + 9, leftStart, "Other", Color.LightGray);

            // Total results
            this.CreateDataCell(topRow + 2, leftStartTotalResults, results.GetOutcomeCount(TestOutcome.Passed, false).ToString());
            this.CreateDataCell(topRow + 3, leftStartTotalResults, results.GetOutcomeCount(TestOutcome.Failed, false).ToString());
            this.CreateDataCell(topRow + 4, leftStartTotalResults, results.GetOutcomeCount(TestOutcome.Blocked, false).ToString());
            this.CreateDataCell(topRow + 5, leftStartTotalResults, results.GetActiveCount(false).ToString());
            this.CreateDataCell(topRow + 6, leftStartTotalResults, results.GetOutcomeCount(TestOutcome.NotExecuted, false).ToString());
            this.CreateDataCell(topRow + 7, leftStartTotalResults, results.GetOutcomeCount(TestOutcome.Inconclusive, false).ToString());
            this.CreateDataCell(topRow + 8, leftStartTotalResults, results.GetOutcomeCount(TestOutcome.Unspecified, false).ToString());
            this.CreateDataCell(topRow + 9, leftStartTotalResults, results.GetOtherCount(false).ToString());

            string sumRange = string.Format("=SUM({0}{1}:{0}{2})", this.LetterForColumn(leftStart + 1), topRow + 2, topRow + 9);
            this.CreateDataCell(topRow + 1, leftStartTotalResults, sumRange);

            // Latest (without duplicates) results
            this.CreateDataCell(topRow + 2, leftStartLatestResults, results.GetOutcomeCount(TestOutcome.Passed).ToString());
            this.CreateDataCell(topRow + 3, leftStartLatestResults, results.GetOutcomeCount(TestOutcome.Failed).ToString());
            this.CreateDataCell(topRow + 4, leftStartLatestResults, results.GetOutcomeCount(TestOutcome.Blocked).ToString());
            this.CreateDataCell(topRow + 5, leftStartLatestResults, results.GetActiveCount().ToString());
            this.CreateDataCell(topRow + 6, leftStartLatestResults, results.GetOutcomeCount(TestOutcome.NotExecuted).ToString());
            this.CreateDataCell(topRow + 7, leftStartLatestResults, results.GetOutcomeCount(TestOutcome.Inconclusive).ToString());
            this.CreateDataCell(topRow + 8, leftStartLatestResults, results.GetOutcomeCount(TestOutcome.Unspecified).ToString());
            this.CreateDataCell(topRow + 9, leftStartLatestResults, results.GetOtherCount().ToString());

            sumRange = string.Format("=SUM({0}{1}:{0}{2})", this.LetterForColumn(leftStart + 2), topRow + 2, topRow + 9);
            this.CreateDataCell(topRow + 1, leftStartLatestResults, sumRange);
        }

        /// <summary>
        /// Get the letter representation of a column
        /// </summary>
        /// <param name="column">The column number</param>
        /// <returns>The alpha value of the given column</returns>
        private string LetterForColumn(int column)
        {
            // Deal with invalid column numbers
            if (column <= 0)
            {
                throw new ArgumentOutOfRangeException("column", "Column numbers must be 1 or greater.");
            }

            column--;
            string columnString = Convert.ToString((char)('A' + (column % 26)));
            while (column >= 26)
            {
                column = (column / 26) - 1;
                columnString = Convert.ToString((char)('A' + (column % 26))) + columnString;
            }

            return columnString;
        } 

        /// <summary>
        /// Format duplicate results
        /// </summary>
        /// <param name="numberOfResults">The number results in the spreadsheet</param>
        private void FormattingDuplicates(int numberOfResults)
        {
            for (int i = 1; i < numberOfResults + 2; i++)
            {
                string primary = ((Range)this.worksheet.Cells[i, 14]).Text;

                if (primary.Equals("FALSE"))
                {
                    this.worksheet.Rows.Range[string.Format("A{0}", i), string.Format("N{0}", i)].Font.Color = Color.DarkGray;
                    this.worksheet.Rows.Range[string.Format("A{0}", i), string.Format("N{0}", i)].Interior.Color = Color.LightYellow;
                }
            }
        }

        /// <summary>
        /// Create data cell with standard black text and white background header
        /// </summary>
        /// <param name="row">The row number</param>
        /// <param name="col">The column number</param>
        /// <param name="text">The text to add to the specific cell</param>
        private void CreateDataCell(int row, int col, string text)
        {
            this.CreateDataCell(row, col, text, Color.White, Color.Black);
        }

        /// <summary>
        /// Create a data cell with black text and the specified background colors
        /// </summary>
        /// <param name="row">The row number</param>
        /// <param name="col">The column number</param>
        /// <param name="text">The text to add to the specific cell</param>
        /// <param name="interiorColor">The background color for the specific cell</param>
        private void CreateDataCell(int row, int col, string text, Color interiorColor)
        {
            this.CreateDataCell(row, col, text, interiorColor, Color.Black);
        }

        /// <summary>
        /// Create a data cell with the specified text and background colors
        /// </summary>
        /// <param name="row">The row number</param>
        /// <param name="col">The column number</param>
        /// <param name="text">The text to add to the specific cell</param>
        /// <param name="interiorColor">The background color of the specific cell</param>
        /// <param name="fontColor">The text color for the specific cell</param>
        private void CreateDataCell(int row, int col, string text, Color interiorColor, Color fontColor)
        {
            // Set the text
            this.worksheet.Cells[row, col] = text;

            // Set the coloring
            Range range = this.worksheet.Cells[row, col];
            range.Interior.Color = interiorColor;
            range.Font.Color = fontColor;

            // Add borders to the header cell
            range.Borders.Color = System.Drawing.Color.Black.ToArgb();
        }

        /// <summary>
        /// Format the header row
        /// </summary>
        private void FormatHeaderRow()
        {
            // Auto fit columns that look funny without auto fit
            this.worksheet.Cells[1, 1].EntireColumn.AutoFit();
            this.worksheet.Cells[1, 9].EntireColumn.AutoFit();
            this.worksheet.Cells[1, 10].EntireColumn.AutoFit();
            this.worksheet.Cells[1, 16].EntireColumn.AutoFit();

            // Format the first row
            Range firstRow = this.worksheet.Rows.Range[string.Format("{0}{1}", this.LetterForColumn(1), 1), string.Format("{0}{1}", this.LetterForColumn(15), 1)];
            this.worksheet.Application.ActiveWindow.SplitRow = 1;
            this.worksheet.Application.ActiveWindow.FreezePanes = true;
            firstRow.Borders.Color = Color.Black;
            firstRow.Interior.Color = Color.DarkKhaki;
            firstRow.AutoFilter(1, Type.Missing, XlAutoFilterOperator.xlAnd, Type.Missing, true);
        }
    }
}