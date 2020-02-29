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

            public UserInterfaceEvents m_uiEvents
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
            public class UI_Button
            {
                private ButtonDefinition _bd;

                public ButtonDefinition bd
                {
                    [MethodImpl(MethodImplOptions.Synchronized)]
                    get
                    {
                        return this._bd;
                    }

                    [MethodImpl(MethodImplOptions.Synchronized)]
                    set
                    {
                        if (this._bd != null)
                        {
                            this._bd.OnExecute -= bd_OnExecute;
                        }

                        this._bd = value;
                        if (this._bd != null)
                        {
                            this._bd.OnExecute += bd_OnExecute;
                        }
                    }
                }

                private void bd_OnExecute(NameValueMap Context)
                {
                    // Link button clicks to their respective commands.
                    switch (bd.InternalName)
                    {
                        case "my_first_button":
                            CommandFunctions.RunAnExe();
                            return;
                        case "my_second_button":
                            CommandFunctions.PopupMessage();
                            return;
                        case "close_doc_button":
                            CommandFunctions.CloseDocument();
                            return;
                        case "export_dxf_button":
                            CommandFunctions.ExportDxf();
                            return;
                        default:
                            return;
                    }
                }
            }

            public delegate ButtonDefinition CreateButton(string display_text, string internal_name, string icon_path);
            public ButtonDefinition button_template(string display_text, string internal_name, string icon_path)
            {
                UI_Button MyButton = new UI_Button();
                MyButton.bd = Utilities.CreateButtonDefinition(display_text, internal_name, "", icon_path);
                return MyButton.bd;
            }

            // Declare all buttons here
            ButtonDefinition MyFirstButton;
            ButtonDefinition MySecondButton;
            ButtonDefinition CloseDocButton;
            ButtonDefinition ExportDxfButton;

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

                    // ButtonName = create_button(display_text, internal_name, icon_path)
                    CreateButton create_button = new CreateButton (button_template);
                    MyFirstButton = create_button("    My First    \n    Command    ", "my_first_button", @"ButtonResources\MyIcon1");
                    MySecondButton = create_button("    My Second    \n    Command    ", "my_second_button", @"ButtonResources\MyIcon2");
                    CloseDocButton = create_button("    Close    \n    Document    ", "close_doc_button", @"ButtonResources\MyIcon3");
                    ExportDxfButton = create_button("    Export    \n    DXF    ", "export_dxf_button", @"ButtonResources\MyIcon4");

                    // Add to the user interface, if it's the first time.
                    // If this add-in doesn't have a UI but runs in the background listening
                    // to events, you can delete this.
                    if (firstTime)
                    {
                        AddToUserInterface();
                    }
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
                CloseDocButton = null;
                ExportDxfButton = null;
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
                Ribbon asmRibbon = Globals.invApp.UserInterfaceManager.Ribbons["Assembly"];
                Ribbon prtRibbon = Globals.invApp.UserInterfaceManager.Ribbons["Part"];
                Ribbon dwgRibbon = Globals.invApp.UserInterfaceManager.Ribbons["Drawing"];


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

                RibbonPanel ExportPanel_prt;
                ExportPanel_prt = setup_panel("Export", "export_panel_prt", MyTab_prt);

                RibbonPanel MyPanel_dwg;
                MyPanel_dwg = setup_panel("My Panel", "my_panel_dwg", MyTab_dwg);

                RibbonPanel MyPanel_asm;
                MyPanel_asm = setup_panel("My Panel", "my_panel_asm", MyTab_asm);


                // Set up Buttons.
                if (!(MyFirstButton == null))
                {
                    MyPanel_asm.CommandControls.AddButton(MyFirstButton, true);
                    MyPanel_dwg.CommandControls.AddButton(MyFirstButton, true);
                }

                if (!(MySecondButton == null))
                {
                    MyPanel_prt.CommandControls.AddButton(MySecondButton, true);
                    MyPanel_dwg.CommandControls.AddButton(MySecondButton, true);
                }

                if (!(CloseDocButton == null))
                {
                    MyPanel_prt.CommandControls.AddButton(CloseDocButton, true);
                    MyPanel_dwg.CommandControls.AddButton(CloseDocButton, true);
                }

                if (!(ExportDxfButton == null))
                {
                    ExportPanel_prt.CommandControls.AddButton(ExportDxfButton, true);
                }
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
                {
                    ribbon_tab = inv_ribbon.RibbonTabs.Add(display_name, internal_name, Globals.g_addInClientID);
                }

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
                {
                    ribbon_panel = ribbon_tab.RibbonPanels.Add(display_name, internal_name, Globals.g_addInClientID);
                }

                setup_panelRet = ribbon_panel;
                return setup_panelRet;
            }


            private void m_uiEvents_OnResetRibbonInterface(NameValueMap Context)
            {
                // The ribbon was reset, so add back the add-ins user-interface.
                AddToUserInterface();
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
        public const string g_simpleAddInClientID = "ff0f87f8-457b-449f-bede-49851b210743";
        public const string g_addInClientID = "{" + g_simpleAddInClientID + "}";
    }
}
