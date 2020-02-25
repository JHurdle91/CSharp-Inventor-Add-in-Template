// TODO: This module exists as a convenient location for the code that does the real
// work when a command is executed.  If you're converting VBA macros into add-in 
// commands you can copy the macros here, convert them to CSharp, 
// and change any references to "ThisApplication" to "Globals.invApp".

using System.Diagnostics;
using System.Windows.Forms;
using Inventor;

namespace My_CSharp_AddIn
{
    public static class CommandFunctions
    {
        public static void RunAnExe()
        {
            var proc = new Process();
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
    }
}
