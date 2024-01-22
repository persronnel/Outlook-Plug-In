using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Vbe.Interop;
using System;

namespace UnityAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // Embed VBA macro code during add-in initialization
            EmbedVbaMacro();
        }

        private void EmbedVbaMacro()
        {
            try
            {
                Outlook.Application outlookApp = Globals.ThisAddIn.Application;

                // Get the VBE (Visual Basic for Applications) object
                VBE vbe = outlookApp.VBE;

                // Add a new VBA module
                VBComponent module = vbe.VBProjects[1].VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);

                // Insert your VBA macro code into the module
                module.CodeModule.AddFromString(@"
Sub YourCreateMacro()
    MsgBox ""Hello from Create Macro!""
End Sub

Sub YourResolveMacro()
    MsgBox ""Hello from Resolve Macro!""
End Sub

Sub YourCloseMacro()
    MsgBox ""Hello from Close Macro!""
End Sub

Sub YourSearchMacro()
    MsgBox ""Hello from Search Macro!""
End Sub

Sub YourEditMacro()
    MsgBox ""Hello from Edit Macro!""
End Sub");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error embedding VBA macro code: {ex.Message}");
            }
        }

        private void create_Click(object sender, RibbonControlEventArgs e)
        {
            // Call the "Create" macro
            RunMacro("YourCreateMacro");
        }

        private void resolve_Click(object sender, RibbonControlEventArgs e)
        {
            // Call the "Resolve" macro
            RunMacro("YourResolveMacro");
        }

        private void close_Click(object sender, RibbonControlEventArgs e)
        {
            // Call the "Close" macro
            RunMacro("YourCloseMacro");
        }

        private void search_Click(object sender, RibbonControlEventArgs e)
        {
            // Call the "Search" macro
            RunMacro("YourSearchMacro");
        }

        private void edt_Click(object sender, RibbonControlEventArgs e)
        {
            // Call the "Edit" macro
            RunMacro("YourEditMacro");
        }

        private void RunMacro(string macroName)
        {
            try
            {
                Outlook.Application outlookApp = Globals.ThisAddIn.Application;

                // Run the specified macro
                outlookApp.Run(macroName);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error running macro '{macroName}': {ex.Message}");
            }
        }
    }
}
