using System;
using Extensibility;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.CommandBars;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Xml;
using System.Threading;
using System.Reflection;
using System.Collections;

namespace BackUpMyData
{
	/// <summary>The object for implementing an Add-in.</summary>
	/// <seealso class='IDTExtensibility2' />
	public class Connect : IDTExtensibility2
	{
		/// <summary>Implements the constructor for the Add-in object. Place your initialization code within this method.</summary>
		public Connect()
		{
		}

		/// <summary>Implements the OnConnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being loaded.</summary>
		/// <param term='application'>Root object of the host application.</param>
		/// <param term='connectMode'>Describes how the Add-in is being loaded.</param>
		/// <param term='addInInst'>Object representing this Add-in.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
		{
			_applicationObject = (DTE2)application;
			_addInInstance = (AddIn)addInInst;
             if (
                           connectMode == ext_ConnectMode.ext_cm_UISetup //The add-in was loaded for user interface setup
                        || connectMode == ext_ConnectMode.ext_cm_AfterStartup //The add-in was loaded when Visual Studio started. 
                        || connectMode == ext_ConnectMode.ext_cm_Startup //The add-in was loaded after Visual Studio started. 
              )
            {
                try
                {
                    CreateContextMenu(_applicationObject);
                }
                catch (Exception)
                {
                    //Most of the error will be due to duplicate menu
                    //We can ignore that errors
                }
            }
		}

        private void CreateContextMenu(DTE2 _applicationObject)
        {
            //Add the main menu
            Commands2 commands = (Commands2)_applicationObject.Commands;
            String[] menus = { "Project and Solution Context Menus" };

            CommandBars contextmenus = (CommandBars)_applicationObject.CommandBars;

            //Add the main menu

            foreach (CommandBar contextmenu in contextmenus)
            {
                if (contextmenu.Name.Contains("Solution"))
                {


                    
                    CommandBarButton oControl = (CommandBarButton)
                    contextmenu.Controls.Add(MsoControlType.msoControlButton,
                    System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value, 1, true);
                    oControl.Caption = "Take BackUp";
                    oControl.Click += oControl_Click;

                    CommandBarButton oImport = (CommandBarButton)
                    contextmenu.Controls.Add(MsoControlType.msoControlButton,
                    System.Reflection.Missing.Value,
                    System.Reflection.Missing.Value, 2, true);
                    oImport.Caption = "Import BackUp..";
                    oImport.Click += oImport_Click;
                    
                     

                }
            }
        }

        void oImport_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            
            OpenFileDialog op = new OpenFileDialog();
            op.Title = "Choose Solution to load back solution at your current Position";
            op.InitialDirectory = Path.GetDirectoryName(Environment.ExpandEnvironmentVariables("%UserProfile%\\Documents\\Visual Studio 2012\\MyBackUpData\\"));
            op.Filter = "sln files (*.sln)|*.sln|All files (*.*)|*.*";

            if (op.ShowDialog() == DialogResult.OK)
            {
                // Backup your Current Solution
                string destlocation = "";
                StringBuilder dest = new StringBuilder();
                string s = Path.GetFileName(_applicationObject.Solution.FullName);
                string solname = _applicationObject.Solution.FullName;
                dest.Append("%UserProfile%\\Documents\\Visual Studio 2012\\MyBackUpData\\");
                dest.Append(s.Substring(0, s.Length - 4) + "_" + DateTime.Today.Year + "_" + DateTime.Today.Month + "_" + DateTime.Today.Day + "_" + DateTime.Now.Hour + "_" + DateTime.Now.Minute + "_" + DateTime.Now.Second);
                destlocation = Environment.ExpandEnvironmentVariables(dest.ToString());
                String orig = Path.GetDirectoryName(_applicationObject.Solution.FullName);
                System.Diagnostics.Process p1 = System.Diagnostics.Process.Start(Environment.ExpandEnvironmentVariables("%Windir%\\System32\\Robocopy.exe"), "\"" + orig + "\"" + " " + "\"" + destlocation + "\"" + " " + " /E /R:0");
                p1.WaitForExit();
                string source = Path.GetDirectoryName(op.FileName);
                System.Diagnostics.Process p = System.Diagnostics.Process.Start(Environment.ExpandEnvironmentVariables("%Windir%\\System32\\Robocopy.exe"), "\"" + source + "\"" + " " + "\"" + orig + "\"" + " " + " /E /R:0");
                p.WaitForExit();
                _applicationObject.Solution.Close(true);

            }
        }



        void oControl_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            StringBuilder source = new StringBuilder();
            StringBuilder dest = new StringBuilder();
            String destlocation = "";
            String[] Value = _applicationObject.Solution.FullName.Split('\\');
            int t=0; 
            
            foreach (string s in Value)
            {
                if (t != (Value.Length - 1))
                {
                    source.Append(s);
                    if(t!=Value.Length - 2)
                    source.Append("\\");
                    t++;
                }
                else
                { 
                    dest.Append("%UserProfile%\\Documents\\Visual Studio 2012\\MyBackUpData\\");
                    dest.Append(s.Substring(0, s.Length - 4) + "_" + DateTime.Today.Year + "_" + DateTime.Today.Month + "_" + DateTime.Today.Day + "_" + DateTime.Now.Hour + "_" + DateTime.Now.Minute + "_" + DateTime.Now.Second);
                    destlocation=Environment.ExpandEnvironmentVariables(dest.ToString());
                }
            }
            System.Diagnostics.Process.Start(Environment.ExpandEnvironmentVariables("%Windir%\\System32\\Robocopy.exe"), "\"" + source.ToString() + "\"" + " " + "\"" + destlocation + "\"" + " " + " /E /R:0");
         
        }



		/// <summary>Implements the OnDisconnection method of the IDTExtensibility2 interface. Receives notification that the Add-in is being unloaded.</summary>
		/// <param term='disconnectMode'>Describes how the Add-in is being unloaded.</param>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom)
		{
		}

		/// <summary>Implements the OnAddInsUpdate method of the IDTExtensibility2 interface. Receives notification when the collection of Add-ins has changed.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />		
		public void OnAddInsUpdate(ref Array custom)
		{
		}

		/// <summary>Implements the OnStartupComplete method of the IDTExtensibility2 interface. Receives notification that the host application has completed loading.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnStartupComplete(ref Array custom)
		{
		}

		/// <summary>Implements the OnBeginShutdown method of the IDTExtensibility2 interface. Receives notification that the host application is being unloaded.</summary>
		/// <param term='custom'>Array of parameters that are host application specific.</param>
		/// <seealso class='IDTExtensibility2' />
		public void OnBeginShutdown(ref Array custom)
		{
		}



		private DTE2 _applicationObject;
		private AddIn _addInInstance;
	}
}