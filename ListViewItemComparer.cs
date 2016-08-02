using System.Collections;
using System.Windows.Forms;

namespace TestPlanViewer
{
    /// <summary>
    /// ListViewItemComparer object
    /// </summary>
    public class ListViewItemComparer : IComparer
    {
        /// <summary>
        /// The column we are sorting on
        /// </summary>
        private int column;

        /// <summary>
        /// The sort order
        /// </summary>
        private SortOrder order;

        /// <summary>
        /// Initializes a new instance of the ListViewItemComparer class
        /// </summary>
        public ListViewItemComparer()
        {
            this.column = 0;
            this.order = SortOrder.Ascending;
        }

        /// <summary>
        /// Initializes a new instance of the ListViewItemComparer class
        /// </summary>
        /// <param name="column">The column to sort on</param>
        /// <param name="order">What order to sort with</param>
        public ListViewItemComparer(int column, SortOrder order)
        {
            this.column = column;
            this.order = order;
        }

        /// <summary>
        /// Compare two list view items
        /// </summary>
        /// <param name="x">The first item</param>
        /// <param name="y">The second item</param>
        /// <returns>The compare int value</returns>
        public int Compare(object x, object y)
        {
            int returnVal = string.Compare(((ListViewItem)x).SubItems[this.column].Text, ((ListViewItem)y).SubItems[this.column].Text);

            // Reverse order for decending
            if (this.order == SortOrder.Descending)
            {
                returnVal *= -1;
            }

            return returnVal;
        }
    }
}
