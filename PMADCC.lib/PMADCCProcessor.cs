/// <summary>
///   Solution : PMADCC
///   Project : PMADCC.lib.dll
///   Module : PMADCCProcessor.cs
///   Description :  Library for PMADCC files processing
/// </summary>
///

using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SHDocVw;
using System.Runtime.InteropServices;
using System.Reflection;

namespace PMADCC.Library
{
    /// <summary>
    /// Class for processing of PMDACC files
    /// </summary>
    public static class PMADCCProcessor
    {
        #region DLL Import

        // Free PMADCC content
        [DllImport("PMADCC.Base.lib.dll", EntryPoint = "Init")]
        private static extern bool wrapped_Init();

        // Free PMADCC content
        [DllImport("PMADCC.Base.lib.dll", EntryPoint = "FreePMADCCData")]
        private static extern IntPtr wrapped_FreePMADCCData(IntPtr data);
        
        // Load PMADCC file
        [DllImport("PMADCC.Base.lib.dll", CharSet = CharSet.Ansi, EntryPoint = "LoadPMADCCDiagram")]
        private static extern bool wrapped_LoadPMADCCDiagram([MarshalAs(UnmanagedType.LPStr)]string diagramID, [MarshalAs(UnmanagedType.LPStr)]string fileName);

        // Load PMADCC file
        [DllImport("PMADCC.Base.lib.dll", CharSet = CharSet.Ansi, EntryPoint = "LoadPMADCCDiagramContent")]
        private static extern IntPtr wrapped_LoadPMADCCDiagramContent([MarshalAs(UnmanagedType.LPStr)]string diagramID, [MarshalAs(UnmanagedType.LPStr)]string fileName, ref ulong resultSize);

        // Save PMADCC file
        [DllImport("PMADCC.Base.lib.dll", CharSet = CharSet.Ansi, EntryPoint = "SavePMADCCDiagram")]
        private static extern bool wrapped_SavePMADCCDiagram([MarshalAs(UnmanagedType.LPStr)]string diagramID);

        // Save PMADCC file
        [DllImport("PMADCC.Base.lib.dll", CharSet = CharSet.Ansi, EntryPoint = "SavePMADCCDiagramAs")]
        private static extern bool wrapped_SavePMADCCDiagramAs([MarshalAs(UnmanagedType.LPStr)]string diagramID, [MarshalAs(UnmanagedType.LPStr)]string fileName);

        // Save PMADCC file
        [DllImport("PMADCC.Base.lib.dll", CharSet = CharSet.Ansi, EntryPoint = "SavePMADCCDiagramContent")]
        private static extern bool wrapped_SavePMADCCDiagramContent([MarshalAs(UnmanagedType.LPStr)]string diagramID, [MarshalAs(UnmanagedType.LPStr)]string fileName, [MarshalAs(UnmanagedType.LPStr)]string svgContent);

        // Reset PMADCC diagram
        [DllImport("PMADCC.Base.lib.dll", CharSet = CharSet.Ansi, EntryPoint = "ResetPMADCCDiagram")]
        private static extern bool wrapped_ResetPMADCCDiagram([MarshalAs(UnmanagedType.LPStr)]string diagramID);

        // Show PMADCC diagram in Internet Explorer
        [DllImport("PMADCC.Base.lib.dll", CharSet = CharSet.Ansi, EntryPoint = "ShowPMADCCDiagram")]
        private static extern bool wrapped_ShowPMADCCDiagram([MarshalAs(UnmanagedType.LPStr)]string diagramID, bool original);
        
        // Get type of object by ID
        [DllImport("PMADCC.Base.lib.dll", CharSet = CharSet.Ansi, EntryPoint = "GetObjectType")]
        private static extern IntPtr wrapped_GetObjectType([MarshalAs(UnmanagedType.LPStr)]string diagramID, int ID, StringBuilder typeName);

        // Get CSS style of object by ID
        [DllImport("PMADCC.Base.lib.dll", CharSet = CharSet.Ansi, EntryPoint = "GetObjectStyleName")]
        private static extern IntPtr wrapped_GetObjectStyleName([MarshalAs(UnmanagedType.LPStr)]string diagramID, int ID, [MarshalAs(UnmanagedType.LPStr)]string objType, StringBuilder styleName);

        // Set CSS style of object by ID
        [DllImport("PMADCC.Base.lib.dll", CharSet = CharSet.Ansi, EntryPoint = "SetObjectStyleName")]
        private static extern IntPtr wrapped_SetObjectStyleName([MarshalAs(UnmanagedType.LPStr)]string diagramID, int ID, [MarshalAs(UnmanagedType.LPStr)]string objType, StringBuilder styleName);
        
        // Get background color of object by ID
        [DllImport("PMADCC.Base.lib.dll", CharSet = CharSet.Ansi, EntryPoint = "GetObjectBackgroundColor")]
        private static extern uint wrapped_GetObjectBackgroundColor([MarshalAs(UnmanagedType.LPStr)]string diagramID, int ID, [MarshalAs(UnmanagedType.LPStr)]string objType);

        // Set background color of object by ID
        [DllImport("PMADCC.Base.lib.dll", CharSet = CharSet.Ansi, EntryPoint = "SetObjectBackgroundColor")]
        private static extern IntPtr wrapped_SetObjectBackgroundColor([MarshalAs(UnmanagedType.LPStr)]string diagramID, int ID, [MarshalAs(UnmanagedType.LPStr)]string objType, StringBuilder styleName, uint colorValue);

        // Get object text value
        [DllImport("PMADCC.Base.lib.dll", CharSet = CharSet.Ansi, EntryPoint = "GetObjectTextValue")]
        private static extern IntPtr wrapped_GetObjectTextValue([MarshalAs(UnmanagedType.LPStr)]string diagramID, int ID, [MarshalAs(UnmanagedType.LPStr)]string objType, StringBuilder textValue);

        // Set object text value
        [DllImport("PMADCC.Base.lib.dll", CharSet = CharSet.Ansi, EntryPoint = "SetObjectTextValue")]
        private static extern IntPtr wrapped_SetObjectTextValue([MarshalAs(UnmanagedType.LPStr)]string diagramID, int ID, [MarshalAs(UnmanagedType.LPStr)]string objType, StringBuilder textValue);


        #endregion

        #region Utils

        /// <summary>
        /// Initialize library
        /// </summary>
        /// <returns>TRUE</returns>
        public static bool Init()
        {
            return wrapped_Init();
        }

        #endregion
        
        #region File operations

        /// <summary>
        /// Load PMDACC file to string
        /// </summary>
        /// <param name="fileName">file name</param>
        /// <returns>string with content of file</returns>
        public static bool LoadPMADCCDiagram(string diagramID, string fileName)
        {
            try
            {
                return wrapped_LoadPMADCCDiagram(diagramID, fileName);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Load PMDAC file to string
        /// </summary>
        /// <param name="fileName">file name</param>
        /// <returns>string with content of file</returns>
        public static string LoadPMADCCDiagramContent(string diagramID, string fileName)
        {
            try
            {
                ulong size = 0;

                IntPtr pResultString = wrapped_LoadPMADCCDiagramContent(diagramID, fileName, ref size);
                byte[] resultBytes = new byte[size];

                Marshal.Copy(pResultString, resultBytes, 0, (int)size);
                wrapped_FreePMADCCData(pResultString);

                return Encoding.Default.GetString(resultBytes);
            }
            catch
            {
                return String.Empty;
            }
        }

        /// <summary>
        /// Save SVG string to PMDAC file
        /// </summary>
        /// <param name="fileName">file name</param>
        /// <returns>TRUE - if succcess</returns>
        public static bool SavePMADCCDiagram(string diagramID, string fileName)
        {
            try
            {
                return wrapped_SavePMADCCDiagramAs(diagramID, fileName);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Save SVG string to PMDAC file
        /// </summary>
        /// <param name="fileName">file name</param>
        /// <returns>TRUE - if succcess</returns>
        public static bool SavePMADCCDiagram(string diagramID)
        {
            try
            {
                return wrapped_SavePMADCCDiagram(diagramID);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Save SVG string to PMDAC file
        /// </summary>
        /// <param name="fileName">file name</param>
        /// <returns>TRUE - if succcess</returns>
        public static bool SavePMADCCDiagramContent(string diagramID, string fileName, string svgContent)
        {
            try
            {
                return wrapped_SavePMADCCDiagramContent(diagramID, fileName, svgContent);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Reset diagram
        /// </summary>
        /// <param name="diagramID">digram ID</param>
        /// <returns>TRUE - if success</returns>
        public static bool ResetPMADCCDiagram(string diagramID)
        {
            try
            {
                return wrapped_ResetPMADCCDiagram(diagramID);
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Internet Explorer

        /// <summary>
        /// Show SVG in Internet Explorer
        /// </summary>
        /// <returns>TRUE - if success</returns>
        public static bool ShowPMADCCDiagram(string diagramID, bool original = false)
        {
            try
            {
                return wrapped_ShowPMADCCDiagram(diagramID, original);
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region SVG content

        #region Styles

        /// <summary>
        /// Get style name
        /// </summary>
        /// <param name="diagramID">ID of diagram</param>
        /// <param name="ID">element ID</param>
        /// <param name="objType">object type</param>
        /// <returns>style name</returns>
        public static string GetObjectStyleName(string diagramID, int ID, string objType)
        {
            try
            {
                StringBuilder styleName = new StringBuilder(32);
                wrapped_GetObjectStyleName(diagramID, ID, objType, styleName);
                return styleName.ToString();
            }
            catch
            {
                return String.Empty;
            }
        }

        /// <summary>
        /// Set style name
        /// </summary>
        /// <param name="diagramID">ID of diagram</param>
        /// <param name="ID">element ID</param>
        /// <param name="styleName">style name</param>
        /// <param name="outStyleName">previuos style name</param>
        /// <returns>XML string with SVG diagram</returns>
        public static string SetObjectStyleName(string diagramID, int ID, string objType, string styleName)
        {
            try
            {
                StringBuilder styleStringBuilder = new StringBuilder(styleName, styleName.Length * 20);
                wrapped_SetObjectStyleName(diagramID, ID, objType, styleStringBuilder);
                return styleStringBuilder.ToString();
            }
            catch
            {
                return styleName;
            }
        }

        #endregion

        #region Color

        /// <summary>
        /// Set background color of element 
        /// </summary>
        /// <param name="diagramID">ID of diagram</param>
        /// <param name="ID">element ID</param>
        /// <param name="objectType">object type</param>
        /// <param name="color">color value</param>
        /// <returns>updated XML string with SVG diagram</returns>
        public static string SetObjectBackgroundColor(string diagramID, int ID, string objectType, uint color)
        {
            try
            {
                StringBuilder styleStringBuilder = new StringBuilder(32);
                wrapped_SetObjectBackgroundColor(diagramID, ID, objectType, styleStringBuilder, color);
                return styleStringBuilder.ToString();
            }
            catch
            {
                return String.Empty;
            }
        }

        /// <summary>
        /// Get element backgound color 
        /// </summary>
        /// <param name="diagramID">ID of diagram</param>
        /// <param name="ID">element ID</param>
        /// <returns>color value</returns>
        public static uint GetObjectBackgroundColor(string diagramID, int ID, string objType)
        {
            try
            {
                return wrapped_GetObjectBackgroundColor(diagramID, ID, objType);
            }
            catch
            {
                return uint.MaxValue;
            }
        }

        #endregion

        #region Type

        /// <summary>
        /// Get element type
        /// </summary>
        /// <param name="diagramID">ID of diagram</param>
        /// <param name="ID">element ID</param>
        /// <returns>type value</returns>
        public static string GetObjectType(string diagramID, int ID)
        {
            try
            {
                StringBuilder typeName = new StringBuilder(64);
                wrapped_GetObjectType(diagramID, ID, typeName);
                return typeName.ToString();
            }
            catch
            {
                return String.Empty;
            }
        }

        #endregion

        #region Text

        /// <summary>
        /// Get element text
        /// </summary>
        /// <param name="diagramID">ID of diagram</param>
        /// <param name="ID">element ID</param>
        /// <returns>text value</returns>
        public static string GetObjectTextValue(string diagramID, int ID)
        {
            try
            {
                StringBuilder textValue = new StringBuilder(1024);
                wrapped_GetObjectTextValue(diagramID, ID, "pmadcc_text", textValue);
                return textValue.ToString();
            }
            catch
            {
                return String.Empty;
            }
        }

        /// <summary>
        /// Set element text
        /// </summary>
        /// <param name="diagramID">ID of diagram</param>
        /// <param name="ID">element ID</param>
        /// <returns>text value</returns>
        public static string SetObjectTextValue(string diagramID, int ID, string text)
        {
            try
            {
                StringBuilder textValue = new StringBuilder(text);
                wrapped_SetObjectTextValue(diagramID, ID, "pmadcc_text", textValue);
                return textValue.ToString();
            }
            catch
            {
                return String.Empty;
            }
        }

        #endregion

        #endregion
    }
}
