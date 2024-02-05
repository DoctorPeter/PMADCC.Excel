/// <summary>
///   Solution : PMADCC
///   Project : PMADCC.Excel.dll
///   Module : AddinRibbonComponent.cs
///   Description :  Add-In control module
/// </summary>
/// 

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

using PMADCC.Library;

namespace PMADCCExcel
{
    /// <summary>
    /// Add-in control class
    /// </summary>
    public partial class AddinRibbonComponent
    {
        /// <summary>
        /// Load event handler
        /// </summary>
        private void AddinRibbonComponent_Load(object sender, RibbonUIEventArgs e)
        {
            // For future use
        }

        /// <summary>
        /// Click on Import from PMADCC button
        /// </summary>
        private void importPMADCCCommand_Click(object sender, RibbonControlEventArgs e)
        {
            if (openPmadccFileDialog.ShowDialog() == DialogResult.OK)
            {
                ProgressForm progressForm = new ProgressForm();
                progressForm.Show();
                progressForm.Refresh();

                if (!Globals.ThisAddIn.OpenPMADCC(openPmadccFileDialog.FileName))
                {
                    // Show progress form
                    progressForm.progressValue = "Error: Can't open " + openPmadccFileDialog.FileName;
                    progressForm.ShowOKButton();
                    progressForm.ShowDialog();
                }
                else
                    progressForm.Close();
            }
        }

        /// <summary>
        /// Click on Export to PMADCC button
        /// </summary>
        private void exportPMADCCCommand_Click(object sender, RibbonControlEventArgs e)
        {
            // save to PMD file
            if (savePmadccFileDialog.ShowDialog() == DialogResult.OK)
            {
                ProgressForm progressForm = new ProgressForm();
                progressForm.Show();
                progressForm.Refresh();

                if (!Globals.ThisAddIn.SavePMADCC(savePmadccFileDialog.FileName))
                {
                    // Show progress form
                    progressForm.progressValue = "Error: Can't save " + savePmadccFileDialog.FileName;
                    progressForm.ShowOKButton();
                    progressForm.ShowDialog();
                }
                else
                    progressForm.Close();
            }
        }

        /// <summary>
        /// Click on Show Diagram button
        /// </summary>
        private void showDiagramButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisAddIn.ShowDiagram(true))
            {
                // Show progress form
                ProgressForm progressForm = new ProgressForm();
                progressForm.progressValue = "Error: Can't send diagram to Internet Explorer!";
                progressForm.ShowOKButton();
                progressForm.ShowDialog();
            }
        }

        /// <summary>
        /// Click on Show Diagram button
        /// </summary>
        private void showWorkDiagramButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisAddIn.ShowDiagram(false))
            {
                // Show progress form
                ProgressForm progressForm = new ProgressForm();
                progressForm.progressValue = "Error: Can't send diagram to Internet Explorer!";
                progressForm.ShowOKButton();
                progressForm.ShowDialog();
            }
        }

        /// <summary>
        /// Click Reset PMADCC button
        /// </summary>
        private void resetPMADCCCommand_Click(object sender, RibbonControlEventArgs e)
        {
            if (!Globals.ThisAddIn.ResetDiagram())
            {
                // Show progress form
                ProgressForm progressForm = new ProgressForm();
                progressForm.progressValue = "Error: Can't reset diagram!";
                progressForm.ShowOKButton();
                progressForm.ShowDialog();
            }
        }
    }
}
