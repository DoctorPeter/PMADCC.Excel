/// <summary>
///   Solution : PMADCC
///   Project : PMADCC.Excel.UDF.dll
///   Module : PMADCCExcelProcessor.cs
///   Description :  COM library for Excel AddIn
/// </summary>
///

using System;
using Extensibility;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Reflection;
using System.Threading;
using Microsoft.Office.Core;

using PMADCC.Library;

namespace PMADCCExcelFunctions
{
    /// <summary>
    /// PMADCC UDF module interface
    /// </summary>
    [Guid("D88B7469-FEFD-4CC0-A169-681404696EED")]
    public interface IPMADCCExcelUDF
    {
        // Load PMADCC file
        bool pmadcc_load(object diagramID, object fileName);

        // Save PMADCC file
        bool pmadcc_save(object diagramID);

        // Save PMADCC file
        bool pmadcc_save_as(object diagramID, object fileName);

        // Reset PMADCC diagram
        bool pmadcc_reset(object diagramID);

        // Show PMADCC diagram in Internet Explorer
        bool pmadcc_showgraph(object diagramID, object original);

        // Change background color
        string pmadcc_setbgcolor(object diagramID, object ID, object objectType, object color);

        // Change background color of step object
        string pmadcc_setstepbgcolor(object diagramID, object ID, object color);

        // Change background color of text object
        string pmadcc_settextbgcolor(object diagramID, object ID, object color);

        // Get background color
        uint pmadcc_getbgcolor(object diagramID, object ID, object objectType);

        // Get background color of step object
        uint pmadcc_getstepbgcolor(object diagramID, object ID);

        // Get background color of text object
        uint pmadcc_gettextbgcolor(object diagramID, object ID);

        // Get CSS Style name
        string pmadcc_getstyle(object diagramID, object ID, object objectType);

        // Get CSS Style name of step object
        string pmadcc_getstepstyle(object diagramID, object ID);

        // Get CSS Style name of text object
        string pmadcc_gettextstyle(object diagramID, object ID);

        // Set CSS Style name
        string pmadcc_setstyle(object diagramID, object ID, object objectType, object styleName);

        // Set CSS Style name of step object
        string pmadcc_setstepstyle(object diagramID, object ID, object styleName);

        // Set CSS Style name of text object
        string pmadcc_settextstyle(object diagramID, object ID, object styleName);

        // Get object type
        string pmadcc_gettype(object diagramID, object ID);

        // Get object text value
        string pmadcc_gettext(object diagramID, object ID);

        // Set object text value
        string pmadcc_settext(object diagramID, object ID, object text);
    }

    /// <summary>
    /// PMADCC UDF functions implementation
    /// </summary>
    [Guid("F0AF62D5-F017-4965-9142-689EA2DE7B9D"),
     ProgId("PMADCCExcelFunctions.PMADCCExcelUDF"),
     ClassInterface(ClassInterfaceType.AutoDual),
     ComVisible(true),
     ComDefaultInterface(typeof(IPMADCCExcelUDF))]
    public class PMADCCExcelUDF : StandardOleMarshalObject, Extensibility.IDTExtensibility2, IPMADCCExcelUDF
    {
        #region Construction

        /// <summary>
        /// Constructor
        /// </summary>
        public PMADCCExcelUDF()
        {
            // For future use
        }

        #endregion

        #region UDFs

        #region File operations

        /// <summary>
        /// UDF for loading of PMADCC file
        /// </summary>
        /// <param name="diagramID">diagram ID</param>
        /// <param name="fileName">file name</param>
        /// <returns>TRUE - if success</returns>
        public bool pmadcc_load(object diagramID, object fileName)
        {
            try
            {
                string diagramIDValue = String.Empty;

                if (diagramID is Excel.Range)
                    diagramIDValue = ((Excel.Range)diagramID).Value2.ToString();
                else
                    diagramIDValue = diagramID.ToString();
                
                string fileNameValue = String.Empty;

                if (fileName is Excel.Range)
                    fileNameValue = ((Excel.Range)fileName).Value2.ToString();
                else
                    fileNameValue = fileName.ToString();

                return PMADCCProcessor.LoadPMADCCDiagram(diagramIDValue, fileNameValue);
            }
            catch
            {
                return false;
            }
        }


        /// <summary>
        /// UDF for saving of PMADCC file
        /// </summary>
        /// <param name="diagramID">diagram ID</param>
        /// <returns>TRUE - if success</returns>
        public bool pmadcc_save(object diagramID)
        {
            try
            {
                string diagramIDValue = String.Empty;

                if (diagramID is Excel.Range)
                    diagramIDValue = ((Excel.Range)diagramID).Value2.ToString();
                else
                    diagramIDValue = diagramID.ToString();

                return PMADCCProcessor.SavePMADCCDiagram(diagramIDValue);
            }
            catch
            {
                return false;
            }
        }


        /// <summary>
        /// UDF for saving of PMADCC file
        /// </summary>
        /// <param name="diagramID">diagram ID</param>
        /// <param name="fileName">file name</param>
        /// <returns>TRUE - if success</returns>
        public bool pmadcc_save_as(object diagramID, object fileName)
        {
            try
            {
                string diagramIDValue = String.Empty;

                if (diagramID is Excel.Range)
                    diagramIDValue = ((Excel.Range)diagramID).Value2.ToString();
                else
                    diagramIDValue = diagramID.ToString();

                string fileNameValue = String.Empty;

                if (fileName is Excel.Range)
                    fileNameValue = ((Excel.Range)fileName).Value2.ToString();
                else
                    fileNameValue = fileName.ToString();

                return PMADCCProcessor.SavePMADCCDiagram(diagramIDValue, fileNameValue);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Reset diagram
        /// </summary>
        /// <param name="diagramID">diagram ID</param>
        /// <returns>true - if success</returns>
        public bool pmadcc_reset(object diagramID)
        {
            try
            {
                string diagramIDValue = String.Empty;

                if (diagramID is Excel.Range)
                    diagramIDValue = ((Excel.Range)diagramID).Value2.ToString();
                else
                    diagramIDValue = diagramID.ToString();

                return PMADCCProcessor.ResetPMADCCDiagram(diagramIDValue);
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Showing

        /// <summary>
        /// Show diagram in Internet Explorer
        /// </summary>
        /// <returns>TRUE - if success</returns>
        public bool pmadcc_showgraph(object diagramID, object original)
        {
            try
            {
                string diagramIDValue = String.Empty;

                if (diagramID is Excel.Range)
                    diagramIDValue = ((Excel.Range)diagramID).Value2.ToString();
                else
                    diagramIDValue = diagramID.ToString();

                bool originalValue = false;
                
                if (original is Excel.Range)
                    originalValue = Boolean.Parse(((Excel.Range)original).Value2.ToString());
                else
                    originalValue = Boolean.Parse(original.ToString());

                return PMADCCProcessor.ShowPMADCCDiagram(diagramIDValue, originalValue);
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Style
                
        /// <summary>
        /// Get CSS Style name 
        /// </summary>
        /// <param name="ID">object ID</param>
        /// <returns>style name</returns>
        public string pmadcc_getstyle(object diagramID, object ID, object objectType)
        {
            try
            {
                string diagramIDValue = String.Empty;

                if (diagramID is Excel.Range)
                    diagramIDValue = ((Excel.Range)diagramID).Value2.ToString();
                else
                    diagramIDValue = diagramID.ToString();

                int idValue = 0;

                if (ID is Excel.Range)
                    idValue = Int32.Parse(((Excel.Range)ID).Value2.ToString());
                else
                    idValue = Int32.Parse(ID.ToString());

                string objectTypeValue = String.Empty;

                if (objectType is Excel.Range)
                    objectTypeValue = ((Excel.Range)objectType).Value2.ToString();
                else
                    objectTypeValue = objectType.ToString();

                return PMADCCProcessor.GetObjectStyleName(diagramIDValue, idValue, objectTypeValue);
            }
            catch
            {
                return String.Empty;
            }
        }

        /// <summary>
        /// Get CSS Style name of step object
        /// </summary>
        /// <param name="ID">object ID</param>
        /// <returns>style name</returns>
        public string pmadcc_getstepstyle(object diagramID, object ID)
        {
            return pmadcc_getstyle(diagramID, ID, "pmadcc_step");
        }

        /// <summary>
        /// Get CSS Style name of text object
        /// </summary>
        /// <param name="ID">object ID</param>
        /// <returns>style name</returns>
        public string pmadcc_gettextstyle(object diagramID, object ID)
        {
            return pmadcc_getstyle(diagramID, ID, "pmadcc_text");
        }

        /// <summary>
        /// Set CSS Style name 
        /// </summary>
        /// <param name="ID">object ID</param>
        /// <param name="styleName">style name</param>
        /// <returns>previous style name</returns>
        public string pmadcc_setstyle(object diagramID, object ID, object objectType, object styleName)
        {
            try
            {
                string diagramIDValue = String.Empty;

                if (diagramID is Excel.Range)
                    diagramIDValue = ((Excel.Range)diagramID).Value2.ToString();
                else
                    diagramIDValue = diagramID.ToString();

                int idValue = 0;

                if (ID is Excel.Range)
                    idValue = Int32.Parse(((Excel.Range)ID).Value2.ToString());
                else
                    idValue = Int32.Parse(ID.ToString());

                string objectTypeValue = String.Empty;

                if (objectType is Excel.Range)
                    objectTypeValue = ((Excel.Range)objectType).Value2.ToString();
                else
                    objectTypeValue = objectType.ToString();

                string styleNameValue = String.Empty;

                if (styleName is Excel.Range)
                    styleNameValue = ((Excel.Range)styleName).Value2.ToString();
                else
                    styleNameValue = styleName.ToString();

                return PMADCCProcessor.SetObjectStyleName(diagramIDValue, idValue, objectTypeValue, styleNameValue);
            }
            catch
            {
                return String.Empty;
            }
        }

        /// <summary>
        /// Set CSS Style name of step object
        /// </summary>
        /// <param name="ID">object ID</param>
        /// <param name="styleName">style name</param>
        /// <returns>previous style name</returns>
        public string pmadcc_setstepstyle(object diagramID, object ID, object styleName)
        {
            return pmadcc_setstyle(diagramID, ID, "pmadcc_step", styleName);
        }

        /// <summary>
        /// Set CSS Style name of text object
        /// </summary>
        /// <param name="ID">object ID</param>
        /// <param name="styleName">style name</param>
        /// <returns>previous style name</returns>
        public string pmadcc_settextstyle(object diagramID, object ID, object styleName)
        {
            return pmadcc_setstyle(diagramID, ID, "pmadcc_text", styleName);
        }

        #endregion

        #region Color

        /// <summary>
        /// Change background color
        /// </summary>
        /// <param name="ID">object ID</param>
        /// <param name="objectType">object type</param>
        /// <param name="color">color value</param>
        /// <returns>TRUE - if success</returns>
        public string pmadcc_setbgcolor(object diagramID, object ID, object objectType, object color)
        {
            try
            {
                string diagramIDValue = String.Empty;

                if (diagramID is Excel.Range)
                    diagramIDValue = ((Excel.Range)diagramID).Value2.ToString();
                else
                    diagramIDValue = diagramID.ToString();

                int idValue = 0;

                if (ID is Excel.Range)
                    idValue = Int32.Parse(((Excel.Range)ID).Value2.ToString());
                else
                    idValue = Int32.Parse(ID.ToString());

                string objectTypeValue = String.Empty;

                if (objectType is Excel.Range)
                    objectTypeValue = ((Excel.Range)objectType).Value2.ToString();
                else
                    objectTypeValue = objectType.ToString();

                uint colorValue = 0;

                if (color is Excel.Range)
                    colorValue = UInt32.Parse(((Excel.Range)color).Value2.ToString());
                else
                    colorValue = UInt32.Parse(color.ToString());

                return PMADCCProcessor.SetObjectBackgroundColor(diagramIDValue, idValue, objectTypeValue, colorValue);
            }
            catch
            {
                return String.Empty;
            }
        }

        /// <summary>
        /// Change background color of step element
        /// </summary>
        /// <param name="ID">object ID</param>
        /// <param name="color">color value</param>
        /// <returns>TRUE - if success</returns>
        public string pmadcc_setstepbgcolor(object diagramID, object ID, object color)
        {
            return pmadcc_setbgcolor(diagramID, ID, "pmadcc_step", color);
        }

        /// <summary>
        /// Change background color of text element
        /// </summary>
        /// <param name="ID">object ID</param>
        /// <param name="color">color value</param>
        /// <returns>TRUE - if success</returns>
        public string pmadcc_settextbgcolor(object diagramID, object ID, object color)
        {
            return pmadcc_setbgcolor(diagramID, ID, "pmadcc_text", color);
        }


        /// <summary>
        /// Change background color
        /// </summary>
        /// <param name="ID">object ID</param>
        /// <param name="objectType">object type</param>
        /// <returns>TRUE - if success</returns>
        public uint pmadcc_getbgcolor(object diagramID, object ID, object objectType)
        {
            try
            {
                string diagramIDValue = String.Empty;

                if (diagramID is Excel.Range)
                    diagramIDValue = ((Excel.Range)diagramID).Value2.ToString();
                else
                    diagramIDValue = diagramID.ToString();

                int idValue = 0;

                if (ID is Excel.Range)
                    idValue = Int32.Parse(((Excel.Range)ID).Value2.ToString());
                else
                    idValue = Int32.Parse(ID.ToString());

                string objectTypeValue = String.Empty;

                if (objectType is Excel.Range)
                    objectTypeValue = ((Excel.Range)objectType).Value2.ToString();
                else
                    objectTypeValue = objectType.ToString();

                return PMADCCProcessor.GetObjectBackgroundColor(diagramIDValue, idValue, objectTypeValue);
            }
            catch
            {
                return uint.MaxValue;
            }
        }

        /// <summary>
        /// Change background color of step element
        /// </summary>
        /// <param name="ID">object ID</param>
        /// <returns>TRUE - if success</returns>
        public uint pmadcc_getstepbgcolor(object diagramID, object ID)
        {
            return pmadcc_getbgcolor(diagramID, ID, "pmadcc_step");
        }

        /// <summary>
        /// Change background color of text element
        /// </summary>
        /// <param name="ID">object ID</param>
        /// <returns>TRUE - if success</returns>
        public uint pmadcc_gettextbgcolor(object diagramID, object ID)
        {
            return pmadcc_getbgcolor(diagramID, ID, "pmadcc_text");
        }

        #endregion

        #region Text

        /// <summary>
        /// Get text value of object
        /// </summary>
        /// <param name="ID">object ID</param>
        /// <returns>text value</returns>
        public string pmadcc_gettext(object diagramID, object ID)
        {
            try
            {
                string diagramIDValue = String.Empty;

                if (diagramID is Excel.Range)
                    diagramIDValue = ((Excel.Range)diagramID).Value2.ToString();
                else
                    diagramIDValue = diagramID.ToString();

                int idValue = 0;

                if (ID is Excel.Range)
                    idValue = Int32.Parse(((Excel.Range)ID).Value2.ToString());
                else
                    idValue = Int32.Parse(ID.ToString());

                return PMADCCProcessor.GetObjectTextValue(diagramIDValue, idValue);
            }
            catch
            {
                return String.Empty;
            }
        }

        /// <summary>
        /// Set object text value
        /// </summary>
        /// <param name="diagramID">diagram ID</param>
        /// <param name="ID">object ID</param>
        /// <param name="text">text value</param>
        /// <returns>previous text value</returns>
        public string pmadcc_settext(object diagramID, object ID, object text)
        {
            try
            {
                string diagramIDValue = String.Empty;

                if (diagramID is Excel.Range)
                    diagramIDValue = ((Excel.Range)diagramID).Value2.ToString();
                else
                    diagramIDValue = diagramID.ToString();

                int idValue = 0;

                if (ID is Excel.Range)
                    idValue = Int32.Parse(((Excel.Range)ID).Value2.ToString());
                else
                    idValue = Int32.Parse(ID.ToString());

                string textValue = String.Empty;

                if (text is Excel.Range)
                    textValue = ((Excel.Range)text).Value2.ToString();
                else
                    textValue = text.ToString();

                return PMADCCProcessor.SetObjectTextValue(diagramIDValue, idValue, textValue);
            }
            catch
            {
                return String.Empty;
            }
        }


        #endregion

        #region Type

        /// <summary>
        /// Get object type
        /// </summary>
        /// <param name="ID">object ID</param>
        /// <returns>type value</returns>
        public string pmadcc_gettype(object diagramID, object ID)
        {
            try
            {
                string diagramIDValue = String.Empty;

                if (diagramID is Excel.Range)
                    diagramIDValue = ((Excel.Range)diagramID).Value2.ToString();
                else
                    diagramIDValue = diagramID.ToString();

                int idValue = 0;
                
                if (ID is Excel.Range)
                    idValue = Int32.Parse(((Excel.Range)ID).Value2.ToString());
                else
                    idValue = Int32.Parse(ID.ToString());
                                
                return PMADCCProcessor.GetObjectType(diagramIDValue, idValue);
            }
            catch
            {
                return String.Empty;
            }
        }

        #endregion

        #endregion

        #region IDTExtensibility2

        // Reference to Excel
        private static Excel.Application Application = null; 

        // Reference to VSTO Add-In
        private static object ThisAddIn = null;

        // Registaration flag
        private static bool fVstoRegister = false;


        /// <summary>
        /// Call this from VSTO
        /// to register DLL itself and load every time
        /// </summary>
        public void Register()
        {
            fVstoRegister = true;
            RegisterFunction(typeof(PMADCCExcelUDF));
            fVstoRegister = false;
        }

        /// <summary>
        /// Call this from VSTO
        /// to remove DLL itself
        /// </summary>
        public void Unregister()
        {
            fVstoRegister = true;
            UnregisterFunction(typeof(PMADCCExcelUDF));
            fVstoRegister = false;
        }
        /// <summary>
        /// Actions when connect
        /// </summary>
        /// <param name="application">Excel application instance</param>
        /// <param name="connectMode">connection mode</param>
        /// <param name="addInInst">Add-In instance</param>
        /// <param name="custom">custom properties</param>
        public void OnConnection(object application, Extensibility.ext_ConnectMode connectMode, object addInInst, ref System.Array custom)
        {
            try
            {
                // get a reference to the instance of the add-in
                Application = application as Excel.Application;
                ThisAddIn = addInInst;
            }
            catch
            {
            }
        }

        /// <summary>
        /// Actions when disconnect
        /// </summary>
        /// <param name="disconnectMode">disconnect mode</param>
        /// <param name="custom">custom properties</param>
        public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
        {
            try
            {
                // clean up
                Marshal.ReleaseComObject(Application);
                Application = null;
                ThisAddIn = null;
                GC.Collect();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch
            {

            }
        }

        /// <summary>
        /// Actions when update
        /// </summary>
        /// <param name="custom">custom properties</param>
        public void OnAddInsUpdate(ref System.Array custom)
        {
            // For future use
        }

        /// <summary>
        /// Actions when startup complete
        /// </summary>
        /// <param name="custom">custom properties</param>
        public void OnStartupComplete(ref System.Array custom)
        {
            // For future use
        }

        /// <summary>
        /// Actions when begin shutdown
        /// </summary>
        /// <param name="custom">custom properties</param>
        public void OnBeginShutdown(ref System.Array custom)
        {
            // For future use
        }

        /// <summary>
        /// Registers the COM Automation Add-in in the CURRENT USER context
        /// and then registers it in all versions of Excel on the users system
        /// without the need of administrator permissions
        /// </summary>
        [ComRegisterFunction]
        public static void RegisterFunction(Type type)
        {
            string PATH = System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase.Replace("\\", "/");
            string ASSM = Assembly.GetExecutingAssembly().FullName;
            int startPos = ASSM.ToLower().IndexOf("version=") + "version=".Length;
            int len = ASSM.ToLower().IndexOf(",", startPos) - startPos;
            string VER = ASSM.Substring(startPos, len);
            string GUID = "{" + type.GUID.ToString().ToUpper() + "}";
            string NAME = type.Namespace + "." + type.Name;
            string BASE = @"Classes\" + NAME;
            string CLSID = @"Classes\CLSID\" + GUID;

            // Open the key
            RegistryKey CU = Registry.CurrentUser.OpenSubKey("Software", true);

            // Is this version registred?
            RegistryKey key = CU.OpenSubKey(CLSID + @"\InprocServer32\" + VER);
            if (key == null)
            {
                // The version of this class currently being registered DOES NOT
                // exist in the registry - so we will now register it

                // BASE KEY
                // HKEY_CURRENT_USER\CLASSES\{NAME}
                key = CU.CreateSubKey(BASE);
                key.SetValue("", NAME);

                // HKEY_CURRENT_USER\CLASSES\{NAME}\CLSID}
                key = CU.CreateSubKey(BASE + @"\CLSID");
                key.SetValue("", GUID);

                // CLSID
                // HKEY_CURRENT_USER\CLASSES\CLSID\{GUID}
                key = CU.CreateSubKey(CLSID);
                key.SetValue("", NAME);


                // HKEY_CURRENT_USER\CLASSES\CLSID\{GUID}\Implemented Categories
                key = CU.CreateSubKey(CLSID + @"\Implemented Categories").CreateSubKey("{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}");

                // HKEY_CURRENT_USER\CLASSES\CLSID\{GUID}\InProcServer32
                key = CU.CreateSubKey(CLSID + @"\InprocServer32");
                key.SetValue("", System.Environment.SystemDirectory + @"\mscoree.dll");
                key.SetValue("ThreadingModel", "Both");
                key.SetValue("Class", NAME);
                key.SetValue("CodeBase", PATH);
                key.SetValue("Assembly", ASSM);
                key.SetValue("RuntimeVersion", "v4.0.30319");

                // HKEY_CURRENT_USER\CLASSES\CLSID\{GUID}\InProcServer32\{VERSION}
                key = CU.CreateSubKey(CLSID + @"\InprocServer32\" + VER);
                key.SetValue("Class", NAME);
                key.SetValue("CodeBase", PATH);
                key.SetValue("Assembly", ASSM);
                key.SetValue("RuntimeVersion", "v4.0.30319");

                // HKEY_CURRENT_USER\CLASSES\CLSID\{GUID}\ProgId
                key = CU.CreateSubKey(CLSID + @"\ProgId");
                key.SetValue("", NAME);

                // HKEY_CURRENT_USER\CLASSES\CLSID\{GUID}\Progammable
                key = CU.CreateSubKey(CLSID + @"\Programmable");

                // Now register the addin in the addins sub keys for each version of Office
                foreach (string keyName in Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\").GetSubKeyNames())
                {
                    if (IsVersionNum(keyName))
                    {
                        // If the adding is found in the Add-in Manager - remove it
                        key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\" + keyName + @"\Excel\Add-in Manager", true);
                        if (key != null)
                        {
                            //key.SetValue(NAME, "");
                            try
                            {
                                key.DeleteValue(NAME);
                            }
                            catch { }
                        }
                    }
                }

                if (!fVstoRegister)
                {
                    // All done - this just helps to assure REGASM is complete
                    // this is not needed, but is useful for troubleshooting
                    MessageBox.Show("Registered " + NAME + ".");
                }
            }
        }

        /// <summary>
        /// Unregisters the add-in, by removing all the keys
        /// </summary>
        [ComUnregisterFunction]
        public static void UnregisterFunction(Type type)
        {
            string GUID = "{" + type.GUID.ToString().ToUpper() + "}";
            string NAME = type.Namespace + "." + type.Name;
            string BASE = @"Classes\" + NAME;
            string CLSID = @"Classes\CLSID\" + GUID;

            // Open the key
            RegistryKey CU = Registry.CurrentUser.OpenSubKey("Software", true);

            // DELETE BASE KEY
            // HKEY_CURRENT_USER\CLASSES\{NAME}
            try
            {
                CU.DeleteSubKeyTree(BASE);
            }
            catch { }
            // HKEY_CURRENT_USER\CLASSES\{NAME}\CLSID}
            try
            {
                CU.DeleteSubKeyTree(CLSID);
            }
            catch { }

            // Now un-register the addin in the addins sub keys for Office
            // here we just make sure to remove it from allversions of Office
            foreach (string keyName in Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\").GetSubKeyNames())
            {
                if (IsVersionNum(keyName))
                {
                    RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\" + keyName + @"\Excel\Add-in Manager", true);
                    if (key != null)
                    {
                        try
                        {
                            key.DeleteValue(NAME);
                        }
                        catch { }
                    }

                    key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\" + keyName + @"\Excel\Options", true);

                    if (key == null)
                        continue;

                    foreach (string valueName in key.GetValueNames())
                    {
                        if (valueName.StartsWith("OPEN"))
                        {
                            if (key.GetValue(valueName).ToString().Contains(NAME))
                            {
                                try
                                {
                                    key.DeleteValue(valueName);
                                }
                                catch { }
                            }
                        }
                    }
                }
            }

            if (!fVstoRegister)
            {
                MessageBox.Show("Unregistered " + NAME + "!");
            }
        }

        /// <summary>
        /// HELPER FUNCTION
        /// This assists is in determining if the subkey string we are passed
        /// is of the type like:
        ///     8.0
        ///     11.0
        ///     14.0
        ///     15.0
        /// </summary>
        public static bool IsVersionNum(string s)
        {
            int idx = s.IndexOf(".");
            if (idx >= 0 && s.EndsWith("0") && int.Parse(s.Substring(0, idx)) > 0)
                return true;
            else
                return false;
        }

        /// <summary>
        /// Set Value of Excel\Options\OPEN
        /// </summary>
        public void SetOpenValue()
        {
            string NAME = "PMADCCExcelFunctions.PMADCCExcelUDF";

            foreach (string keyName in Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\").GetSubKeyNames())
            {
                if (IsVersionNum(keyName))
                {
                    RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\" + keyName + @"\Excel\Options", true);
                    if (key != null)
                    {
                        string openString = "";

                        foreach (string valueName in key.GetValueNames())
                        {
                            if (valueName.Contains("OPEN"))
                                openString = valueName;
                        }

                        if (openString == "")
                            openString = "OPEN";
                        else
                        {
                            string openNumberString = openString.Substring(4, openString.Length - 4);
                            if (openNumberString == "")
                                openString = "OPEN1";
                            else
                                openString = "OPEN" + (Int32.Parse(openNumberString) + 1).ToString();
                        }

                        key.SetValue(openString, "/A " + "\"" + NAME + "\"");
                    }
                }
            }
        }

        #endregion

    }
}
