using Microsoft.TeamFoundation.TestManagement.Client;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace TestPlanViewer
{
    /// <summary>
    /// UI class for a selection box
    /// </summary>
    public partial class SelectPlanDialog : Form 
    {
        /// <summary>
        /// Initializes a new instance of the SelectPlanDialog class
        /// </summary>
        /// <param name="plans">The list of plan names</param>
        public SelectPlanDialog(List<string> plans)
        {
            this.InitializeComponent();
            this.PlansComboBox.Items.AddRange(plans.ToArray());
        }


        /// <summary>
        /// Initializes a new instance of the SelectPlanDialog class
        /// </summary>
        /// <param name="plans">The list of plan names</param>
        public SelectPlanDialog(List<ITestPlan> plans)
        {
            this.InitializeComponent();
            this.PlansComboBox.Items.AddRange(plans.ToArray());
        }


        /// <summary>
        /// Gets the test plan name
        /// </summary>
        public string Plan { get; private set; }

        /// <summary>
        /// The accept button was clicked
        /// </summary>
        /// <param name="sender">The system send object</param>
        /// <param name="e">The system event</param>
        private void AcceptButton_Click(object sender, EventArgs e)
        {
            // Make sure the user selected a plan
            if (string.IsNullOrEmpty(this.PlansComboBox.Text))
            {
                MessageBox.Show("No plan was selected", "Plan Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                this.Plan = this.PlansComboBox.Text;
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
        }

        /// <summary>
        /// The cancel button was clicked
        /// </summary>
        /// <param name="sender">The system send object</param>
        /// <param name="e">The system event</param>
        private void CancelPlanButton_Click(object sender, EventArgs e)
        {
            this.Plan = string.Empty;
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }
    }
}
