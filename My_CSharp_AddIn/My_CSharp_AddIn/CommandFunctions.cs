// TODO: This module exists as a convenient location for the code that does the real
// work when a command is executed.  If you're converting VBA macros into add-in 
// commands you can copy the macros here, convert them to CSharp, 
// and change any references to "ThisApplication" to "Globals.invApp".

using System;
using System.Diagnostics;
using System.Windows.Forms;
using Inventor;

namespace My_CSharp_AddIn
{
    public static class CommandFunctions
    {
        public static void RunAnExe()
        {
            Process proc = new Process();
            proc = Process.Start(@"C:\path_to\some_file.exe", "");
        }

        public static void PopupMessage()
        {
            MessageBox.Show("This is a message box!");
        }

        public static void CloseDocument()
        {
            Globals.invApp.ActiveDocument.Close(true);
        }

        public static void ExportDxf()
        {
            PartDocument doc = (PartDocument)Globals.invApp.ActiveDocument;

            if (doc.DocumentSubType.DocumentSubTypeID != "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}")
            {
                MessageBox.Show("This can only be run on a Sheet Metal document. Exiting...");
                return;
            }

            SheetMetalComponentDefinition smcd;
            smcd = (SheetMetalComponentDefinition)doc.ComponentDefinition;

            if (!smcd.HasFlatPattern)
            {
                try
                {
                    smcd.Unfold();
                    smcd.FlatPattern.ExitEdit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not unfold part. Create flat pattern and try again.\n\n" + ex.Message);
                }
            }

            const string DXF_OPTIONS = "FLAT PATTERN DXF?AcadVersion=2004" +
                                            "&OuterProfileLayer=IV_INTERIOR_PROFILES" +
                                            "&InvisibleLayers=" +
                                                "IV_TANGENT;" +
                                                "IV_FEATURE_PROFILES_DOWN;" +
                                                "IV_BEND;" +
                                                "IV_BEND_DOWN;" +
                                                "IV_TOOL_CENTER;" +
                                                "IV_TOOL_CENTER_DOWN;" +
                                                "IV_ARC_CENTERS;" +
                                                "IV_FEATURE_PROFILES;" +
                                                "IV_FEATURE_PROFILES_DOWN;" +
                                                "IV_ALTREP_FRONT;" +
                                                "IV_ALTREP_BACK;" +
                                                "IV_ROLL_TANGENT;" +
                                                "IV_ROLL" +
                                            "&SimplifySplines=True" +
                                            "&BendLayerColor=255;255;0";

            string DXF_PATH;
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Select DXF Folder";
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                DXF_PATH = fbd.SelectedPath;
            }
            else
            {
                MessageBox.Show("Must specify DXF path");
                return;
            }

            string DISPLAY_NAME = doc.DisplayName;
            string dxf_filename = string.Format("{0}\\{1}.dxf", DXF_PATH, DISPLAY_NAME);
            try
            {
                smcd.DataIO.WriteDataToFile(DXF_OPTIONS, dxf_filename);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to save DXF");
                return;
            }

            MessageBox.Show("DXF saved to path:\n\n" + DXF_PATH);
        }
    }
}
