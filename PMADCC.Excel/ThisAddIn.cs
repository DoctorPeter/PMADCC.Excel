/// <summary>
///   Solution : PMADCC
///   Project : PMADCC.Excel.dll
///   Module : ThisAddIn.cs
///   Description :  Add-In main module
/// </summary>
/// 

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

using PMADCC.Library;

namespace PMADCCExcel
{
    /// <summary>
    /// Excel PMADCC main class
    /// </summary>
    public partial class ThisAddIn
    {
        #region PMADCC UDF instance

        // Reference to PMADCC udf
        private PMADCCExcelFunctions.PMADCCExcelUDF pmadccExcelUDFRef = null;
        
        // Property of PMADCC UDF instance
        private PMADCCExcelFunctions.PMADCCExcelUDF pmadccExcelUDFInstance
        {
            get
            {
                if (pmadccExcelUDFRef == null)
                    pmadccExcelUDFRef = new PMADCCExcelFunctions.PMADCCExcelUDF();

                return pmadccExcelUDFRef;
            }
        }

        #endregion
                
        #region Request for Com AddIn object
       
        /// <summary>
        /// Get automotion Add-in object
        /// </summary>
        protected override object RequestComAddInAutomationService()
        {
            return pmadccExcelUDFInstance;
        }

        #endregion

        #region Startup/Shutdown

        // PMADCCExcelFunctions.PMADCCExcelProcessor Installation flag
        private bool installationSuccess = true;

        /// <summary>
        /// Startup add-in action
        /// </summary>
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Get the name and GUID from the class
            string NAME = pmadccExcelUDFInstance.GetType().Namespace + "." + pmadccExcelUDFInstance.GetType().Name;
            string GUID = pmadccExcelUDFInstance.GetType().GUID.ToString().ToUpper();

            // Is the add-in already loaded in Excel, but maybe disabled
            // if this is the case - try to re-enable it
            bool fFound = false;
            foreach (Excel.AddIn a in Application.AddIns)
            {
                try
                {
                    if (a.CLSID.Contains(GUID))
                    {
                        fFound = true;
                        if (!a.Installed)
                            a.Installed = true;
                        break;
                    }
                }
                catch { }
            }

            // If we do not see the UDF class in the list of installed addin we need to
            // add it to the collection
            if (!fFound)
            {
                pmadccExcelUDFInstance.Register();

                try
                {
                    Application.AddIns.Add(NAME).Installed = true;
                }
                catch
                {
                    MessageBox.Show("This is the first launch of PMADCC Add-in for Microsoft Excel. PMADCC was built in Excel. In order to gain access to its functions, you need to restart Excel.", "PMADCC Setup", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    installationSuccess = false;
                }

            }

            // Initializa base libary
            PMADCCProcessor.Init();
        }

        /// <summary>
        /// Shutdown add-in action
        /// </summary>
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (!installationSuccess)
                pmadccExcelUDFInstance.SetOpenValue();
        }

        #endregion

        #region PMADCC files

        /// <summary>
        /// Save active page to PMADCC file
        /// </summary>
        /// <param name="fileName">file name string</param>
        /// <returns>true - if success</returns>
        public bool SavePMADCC(string fileName)
        {
            try
            {
                return PMADCCProcessor.SavePMADCCDiagram("PMADCC_EXCEL", fileName);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Open PMD file
        /// </summary>
        /// <param name="fileName">file name</param>
        /// <returns>TRUE - if success</returns>
        public bool OpenPMADCC(string fileName)
        {
            try
            {
                return PMADCCProcessor.LoadPMADCCDiagram("PMADCC_EXCEL", fileName);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Shoe data diagram in Intenet explorer
        /// </summary>
        /// <returns>TRUE - if success</returns>
        public bool ShowDiagram(bool original)
        {
            try
            {
                return PMADCCProcessor.ShowPMADCCDiagram("PMADCC_EXCEL", original);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Reset data diagram
        /// </summary>
        /// <returns>TRUE - if success</returns>
        public bool ResetDiagram()
        {
            try
            {
                return PMADCCProcessor.ResetPMADCCDiagram("PMADCC_EXCEL");
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
