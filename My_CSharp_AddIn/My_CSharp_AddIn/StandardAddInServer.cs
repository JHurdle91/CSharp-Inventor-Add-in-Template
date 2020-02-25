using Microsoft.VisualBasic;
using System;
using Inventor;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Windows.Forms;

namespace My_CSharp_AddIn
{
    namespace My_CSharp_AddIn
    {
        [ProgId("My_CSharp_AddIn.StandardAddInServer")]
        [Guid(Globals.g_simpleAddInClientID)]
        public class StandardAddInServer : ApplicationAddInServer
        {

            // *********************************************************************************
            // * The two declarations below are related to adding buttons to Inventor's UI.
            // * They can be deleted if this add-in doesn't have a UI and only runs in the 
            // * background handling events.
            // *********************************************************************************

            // Declaration of the object for the UserInterfaceEvents to be able to handle
            // if the user resets the ribbon so the button can be added back in.
            private UserInterfaceEvents _m_uiEvents;

            private UserInterfaceEvents m_uiEvents
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _m_uiEvents;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_m_uiEvents != null)
                    {
                        _m_uiEvents.OnResetRibbonInterface -= m_uiEvents_OnResetRibbonInterface;
                    }

                    _m_uiEvents = value;
                    if (_m_uiEvents != null)
                    {
                        _m_uiEvents.OnResetRibbonInterface += m_uiEvents_OnResetRibbonInterface;
                    }
                }
            }

            // Declaration of the button definition with events to handle the click event.
            // For additional commands this declaration along with other sections of code
            // that apply to the button can be duplicated from this example.
            private ButtonDefinition _MyFirstButton;

            private ButtonDefinition MyFirstButton
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _MyFirstButton;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_MyFirstButton != null)
                    {
                        _MyFirstButton.OnExecute -= MyFirstButton_OnExecute;
                    }

                    _MyFirstButton = value;
                    if (_MyFirstButton != null)
                    {
                        _MyFirstButton.OnExecute += MyFirstButton_OnExecute;
                    }
                }
            }

            private ButtonDefinition _MySecondButton;

            private ButtonDefinition MySecondButton
            {
                [MethodImpl(MethodImplOptions.Synchronized)]
                get
                {
                    return _MySecondButton;
                }

                [MethodImpl(MethodImplOptions.Synchronized)]
                set
                {
                    if (_MySecondButton != null)
                    {
                        _MySecondButton.OnExecute -= MySecondButton_OnExecute;
                    }

                    _MySecondButton = value;
                    if (_MySecondButton != null)
                    {
                        _MySecondButton.OnExecute += MySecondButton_OnExecute;
                    }
                }
            }


            // This method is called by Inventor when it loads the AddIn. The AddInSiteObject provides access  
            // to the Inventor Application object. The FirstTime flag indicates if the AddIn is loaded for
            // the first time. However, with the introduction of the ribbon this argument is always true.
            public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
            {
                try
                {
                    // Initialize AddIn members.
                    Globals.invApp = addInSiteObject.Application;

                    // Connect to the user-interface events to handle a ribbon reset.
                    m_uiEvents = Globals.invApp.UserInterfaceManager.UserInterfaceEvents;

                    // *********************************************************************************
                    // * The remaining code in this Sub is all for adding the add-in into Inventor's UI.
                    // * It can be deleted if this add-in doesn't have a UI and only runs in the 
                    // * background handling events.
                    // *********************************************************************************

                    // Create the button definition using the CreateButtonDefinition function to simplify this step.
                    // ButtonName = Utilities.CreateButtonDefinition(display_text, internal_name, "", icon_path)
                    MyFirstButton = Utilities.CreateButtonDefinition("    My First    \n    Command    ", "MyFirstCommand", "", @"ButtonResources\MyIcon1");
                    MySecondButton = Utilities.CreateButtonDefinition("    My Second    \n    Command    ", "MySecondCommand", "", @"ButtonResources\MyIcon2");

                    // Add to the user interface, if it's the first time.
                    // If this add-in doesn't have a UI but runs in the background listening
                    // to events, you can delete this.
                    if (firstTime)
                        AddToUserInterface();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Unexpected failure in the activation of the add-in \"My_CSharp_AddIn\"" + System.Environment.NewLine + System.Environment.NewLine + ex.Message);
                }
            }

            // This method is called by Inventor when the AddIn is unloaded. The AddIn will be
            // unloaded either manually by the user or when the Inventor session is terminated.
            public void Deactivate()
            {
                // Release objects.
                MyFirstButton = null;
                MySecondButton = null;
                m_uiEvents = null;
                Globals.invApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            // This property is provided to allow the AddIn to expose an API of its own to other 
            // programs. Typically, this  would be done by implementing the AddIn's API
            // interface in a class and returning that class object through this property.
            // Typically it's not used, like in this case, and returns Nothing.
            public object Automation
            {
                get
                {
                    return null;
                }
            }

            // Note:this method is now obsolete, you should use the 
            // ControlDefinition functionality for implementing commands.
            public void ExecuteCommand(int commandID)
            {
            }


            // Adds whatever is needed by this add-in to the user-interface.  This is 
            // called when the add-in loaded and also if the user interface is reset.
            private void AddToUserInterface()
            {

                // Get the ribbon. (more buttons can be added to various ribbons within this single addin)
                // Ribbons:
                // ZeroDoc
                // Part
                // Assembly
                // Drawing
                // Presentation
                // iFeatures
                // UnknownDocument
                var asmRibbon = Globals.invApp.UserInterfaceManager.Ribbons["Assembly"];
                var prtRibbon = Globals.invApp.UserInterfaceManager.Ribbons["Part"];
                var dwgRibbon = Globals.invApp.UserInterfaceManager.Ribbons["Drawing"];


                // Set up Tabs.
                // tab = setup_panel(display_name, internal_name, inv_ribbon)
                RibbonTab MyTab_asm;
                MyTab_asm = setup_tab("My Tab", "my_tab_asm", asmRibbon);

                RibbonTab MyTab_prt;
                MyTab_prt = setup_tab("My Tab", "my_tab_prt", prtRibbon);

                RibbonTab MyTab_dwg;
                MyTab_dwg = setup_tab("My Tab", "my_tab_dwg", dwgRibbon);


                // Set up Panels.
                // panel = setup_panel(display_name, internal_name, ribbon_tab)
                RibbonPanel MyPanel_prt;
                MyPanel_prt = setup_panel("My Panel", "my_panel_prt", MyTab_prt);

                RibbonPanel MyPanel_dwg;
                MyPanel_dwg = setup_panel("My Panel", "my_panel_dwg", MyTab_dwg);

                RibbonPanel MyPanel_asm;
                MyPanel_asm = setup_panel("My Panel", "my_panel_asm", MyTab_asm);


                // Set up Buttons.
                if (!(MyFirstButton == null))
                    MyPanel_asm.CommandControls.AddButton(MyFirstButton, true);
                MyPanel_dwg.CommandControls.AddButton(MyFirstButton, true);

                if (!(MySecondButton == null))
                    MyPanel_prt.CommandControls.AddButton(MySecondButton, true);
                MyPanel_dwg.CommandControls.AddButton(MySecondButton, true);
            }


            private RibbonTab setup_tab(string display_name, string internal_name, Ribbon inv_ribbon)
            {
                RibbonTab setup_tabRet = default(RibbonTab);
                RibbonTab ribbon_tab = null;
                try
                {
                    ribbon_tab = inv_ribbon.RibbonTabs[internal_name];
                }
                catch (Exception ex)
                {
                }

                if (ribbon_tab == null)
                    ribbon_tab = inv_ribbon.RibbonTabs.Add(display_name, internal_name, Globals.g_addInClientID);

                setup_tabRet = ribbon_tab;
                return setup_tabRet;
            }


            private RibbonPanel setup_panel(string display_name, string internal_name, RibbonTab ribbon_tab)
            {
                RibbonPanel setup_panelRet = default(RibbonPanel);
                RibbonPanel ribbon_panel = null;
                try
                {
                    ribbon_panel = ribbon_tab.RibbonPanels[internal_name];
                }
                catch (Exception ex)
                {
                }

                if (ribbon_panel == null)
                    ribbon_panel = ribbon_tab.RibbonPanels.Add(display_name, internal_name, Globals.g_addInClientID);

                setup_panelRet = ribbon_panel;
                return setup_panelRet;
            }


            private void m_uiEvents_OnResetRibbonInterface(NameValueMap Context)
            {
                // The ribbon was reset, so add back the add-ins user-interface.
                AddToUserInterface();
            }


            // Link button clicks to their respective commands.
            private void MyFirstButton_OnExecute(NameValueMap Context)
            {
                CommandFunctions.RunAnExe();
            }

            private void MySecondButton_OnExecute(NameValueMap Context)
            {
                CommandFunctions.PopupMessage();
            }
        }
    }


    public static class Globals
    {
        // Inventor application object.
        public static Inventor.Application invApp;

        // The unique ID for this add-in.  If this add-in is copied to create a new add-in
        // you need to update this ID along with the ID in the .manifest file, the .addin file
        // and create a new ID for the typelib GUID in AssemblyInfo.vb
        public const string g_simpleAddInClientID = "ff0f87f8-346a-449f-bede-49851b210743";
        public const string g_addInClientID = "{" + g_simpleAddInClientID + "}";
    }
}
