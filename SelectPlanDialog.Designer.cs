﻿// <auto-generated />
namespace TestPlanViewer
{
    partial class SelectPlanDialog
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.PlansComboBox = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.AcceptPlanButton = new System.Windows.Forms.Button();
            this.CancelPlanButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // PlansComboBox
            // 
            this.PlansComboBox.FormattingEnabled = true;
            this.PlansComboBox.Location = new System.Drawing.Point(9, 25);
            this.PlansComboBox.Name = "PlansComboBox";
            this.PlansComboBox.Size = new System.Drawing.Size(260, 21);
            this.PlansComboBox.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(31, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Plan:";
            // 
            // AcceptPlanButton
            // 
            this.AcceptPlanButton.Location = new System.Drawing.Point(113, 52);
            this.AcceptPlanButton.Name = "AcceptPlanButton";
            this.AcceptPlanButton.Size = new System.Drawing.Size(75, 23);
            this.AcceptPlanButton.TabIndex = 2;
            this.AcceptPlanButton.Text = "Accept";
            this.AcceptPlanButton.UseVisualStyleBackColor = true;
            this.AcceptPlanButton.Click += new System.EventHandler(this.AcceptButton_Click);
            // 
            // CancelPlanButton
            // 
            this.CancelPlanButton.Location = new System.Drawing.Point(194, 52);
            this.CancelPlanButton.Name = "CancelPlanButton";
            this.CancelPlanButton.Size = new System.Drawing.Size(75, 23);
            this.CancelPlanButton.TabIndex = 3;
            this.CancelPlanButton.Text = "Cancel";
            this.CancelPlanButton.UseVisualStyleBackColor = true;
            this.CancelPlanButton.Click += new System.EventHandler(this.CancelPlanButton_Click);
            // 
            // SelectPlanDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(276, 81);
            this.ControlBox = false;
            this.Controls.Add(this.CancelPlanButton);
            this.Controls.Add(this.AcceptPlanButton);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.PlansComboBox);
            this.Name = "SelectPlanDialog";
            this.Text = "Select Test Plan";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox PlansComboBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button AcceptPlanButton;
        private System.Windows.Forms.Button CancelPlanButton;
    }
}