/// <summary>
///   Solution : PMADCC
///   Project : PMADCC.Visio.dll
///   Module : ProgressForm.cs
///   Description :  Progress form module
/// </summary>
/// 

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PMADCC.Library
{
    /// <summary>
    /// Progress form class
    /// </summary>
    public partial class ProgressForm : Form
    {

        /// <summary>
        /// Constructor
        /// </summary>
        public ProgressForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// OK button click
        /// </summary>
        private void okButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        // Property of progress text
        public string progressValue
        {
            get
            {
                return progressLabel.Text;
            }

            set
            {
                progressLabel.Text = value;
            }
        }

        /// <summary>
        /// Show OK button
        /// </summary>
        public void ShowOKButton()
        {
            okButton.Visible = true;
        }
    }
}
