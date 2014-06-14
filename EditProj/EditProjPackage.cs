using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using System.IO;
using Microsoft.VisualStudio.Package;
using IOleServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;
using Microsoft.VisualStudio.TextManager.Interop;
using EnvDTE;
using System.Collections.Generic;

namespace rdomunozcom.EditProj
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    ///
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the 
    /// IVsPackage interface and uses the registration attributes defined in the framework to 
    /// register itself and its components with the shell.
    /// </summary>
    // This attribute tells the PkgDef creation utility (CreatePkgDef.exe) that this class is
    // a package.
    [PackageRegistration(UseManagedResourcesOnly = true)]
    // This attribute is used to register the information needed to show this package
    // in the Help/About dialog of Visual Studio.
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    // This attribute is needed to let the shell know that this package exposes some menus.
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidEditProjPkgString)]
    public sealed class EditProjPackage : Package
    {

        private CommandEvents saveFileCommand, saveAllCommand, exitCommand;
        private DTE dte;
        private IDictionary<string, string> tempToProjFiles;

        /// <summary>
        /// Default constructor of the package.
        /// Inside this method you can place any initialization code that does not require 
        /// any Visual Studio service because at this point the package object is created but 
        /// not sited yet inside Visual Studio environment. The place to do all the other 
        /// initialization is the Initialize method.
        /// </summary>
        public EditProjPackage()
        {
            this.tempToProjFiles = new Dictionary<string, string>();
        }



        /////////////////////////////////////////////////////////////////////////////
        // Overridden Package Implementation
        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            Debug.WriteLine(string.Format(CultureInfo.CurrentCulture, "Entering Initialize() of: {0}", this.ToString()));
            base.Initialize();

            // Add our command handlers for menu (commands must exist in the .vsct file)
            OleMenuCommandService mcs = this.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (null != mcs)
            {
                // Create the command for the menu item.
                CommandID menuCommandID = new CommandID(GuidList.guidEditProjCmdSet, (int)PkgCmdIDList.editProjFile);
                var menuItem = new OleMenuCommand(MenuItemCallback, menuCommandID);
                menuItem.BeforeQueryStatus += menuCommand_BeforeQueryStatus;

                this.dte = this.GetService(typeof(DTE)) as DTE;
                // need to keep a strong reference to CommandEvents so that it's not GC'ed
                this.saveFileCommand = this.dte.Events.CommandEvents[VSConstants.CMDSETID.StandardCommandSet97_string, (int)VSConstants.VSStd97CmdID.SaveProjectItem];
                this.saveAllCommand = this.dte.Events.CommandEvents[VSConstants.CMDSETID.StandardCommandSet97_string, (int)VSConstants.VSStd97CmdID.SaveSolution];

                this.saveFileCommand.AfterExecute += saveCommands_AfterExecute;
                this.saveAllCommand.AfterExecute += saveCommands_AfterExecute;
                mcs.AddCommand(menuItem);
            }
        }

        private void menuCommand_BeforeQueryStatus(object sender, EventArgs e)
        {
            // get the menu that fired the event
            var menuCommand = sender as OleMenuCommand;
            if (menuCommand != null)
            {
                // start by assuming that the menu will not be shown
                menuCommand.Visible = false;
                menuCommand.Enabled = false;

                IVsHierarchy hierarchy = null;
                uint itemid = VSConstants.VSITEMID_NIL;

                //if (!IsSingleProjectItemSelection(out hierarchy, out itemid)) return;
                //// Get the file path
                //string itemFullPath = null;
                //((IVsProject)hierarchy).GetMkDocument(itemid, out itemFullPath);
                //var transformFileInfo = new FileInfo(itemFullPath);

                // then check if the file is named 'web.config'
                //bool isWebConfig = string.Compare("web.config", transformFileInfo.Name, StringComparison.OrdinalIgnoreCase) == 0;

                //// if not leave the menu hidden
                //if (!isWebConfig) return;

                menuCommand.Visible = true;
                menuCommand.Enabled = true;
            }
        }

        private static bool IsSingleProjectItemSelection(out IVsHierarchy hierarchy, out uint itemid)
        {
            hierarchy = null;
            itemid = VSConstants.VSITEMID_NIL;
            int hr = VSConstants.S_OK;

            IVsMonitorSelection monitorSelection = Package.GetGlobalService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            var solution = Package.GetGlobalService(typeof(SVsSolution)) as IVsSolution;
            if (monitorSelection == null || solution == null)
            {
                return false;
            }

            IVsMultiItemSelect multiItemSelect = null;
            IntPtr hierarchyPtr = IntPtr.Zero;
            IntPtr selectionContainerPtr = IntPtr.Zero;

            try
            {
                hr = monitorSelection.GetCurrentSelection(out hierarchyPtr, out itemid, out multiItemSelect, out selectionContainerPtr);

                if (ErrorHandler.Failed(hr) || hierarchyPtr == IntPtr.Zero || itemid == VSConstants.VSITEMID_NIL)
                {
                    // there is no selection
                    return false;
                }

                // multiple items are selected
                if (multiItemSelect != null) return false;

                // there is a hierarchy root node selected, thus it is not a single item inside a project

                if (itemid == VSConstants.VSITEMID_ROOT) return true;

                hierarchy = Marshal.GetObjectForIUnknown(hierarchyPtr) as IVsHierarchy;
                if (hierarchy == null) return false;

                Guid guidProjectID = Guid.Empty;

                if (ErrorHandler.Failed(solution.GetGuidOfProject(hierarchy, out guidProjectID)))
                {
                    return false; // hierarchy is not a project inside the Solution if it does not have a ProjectID Guid
                }

                // if we got this far then there is a single project item selected
                return true;
            }
            finally
            {
                if (selectionContainerPtr != IntPtr.Zero)
                {
                    Marshal.Release(selectionContainerPtr);
                }

                if (hierarchyPtr != IntPtr.Zero)
                {
                    Marshal.Release(hierarchyPtr);
                }
            }
        }
        #endregion

        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            IVsUIShellOpenDocument openDocument = Package.GetGlobalService(typeof(SVsUIShellOpenDocument)) as IVsUIShellOpenDocument;

            IVsMonitorSelection monitorSelection = Package.GetGlobalService(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;
            IOleServiceProvider serviceProvider = GetService(typeof(IOleServiceProvider)) as IOleServiceProvider;

            IVsMultiItemSelect multiItemSelect = null;
            IntPtr hierarchyPtr = IntPtr.Zero;
            IntPtr selectionContainerPtr = IntPtr.Zero;
            int hr = VSConstants.S_OK;
            uint itemid = VSConstants.VSITEMID_NIL;

            hr = monitorSelection.GetCurrentSelection(out hierarchyPtr, out itemid, out multiItemSelect, out selectionContainerPtr);

            IVsHierarchy hierarchy = Marshal.GetObjectForIUnknown(hierarchyPtr) as IVsHierarchy;
            IVsUIHierarchy uiHierarchy = hierarchy as IVsUIHierarchy;

            // Get the file path
            string projFilePath = null;
            ((IVsProject)hierarchy).GetMkDocument(itemid, out projFilePath);

            string projectData = File.ReadAllText(projFilePath);
            string tempDir = Path.GetTempPath();
            string tempProjFile = Guid.NewGuid().ToString() + ".xml";
            string tempProjFilePath = Path.Combine(tempDir, tempProjFile);
            this.tempToProjFiles[tempProjFilePath] = projFilePath;
            File.WriteAllText(tempProjFilePath, projectData);

            OpenFileInEditor(Path.Combine(tempDir, tempProjFilePath), Resources.Project, uiHierarchy, itemid);
        }


        private void OpenFileInEditor(string filePath, string fileName, IVsUIHierarchy uiHierarchy, UInt32 itemID, bool openWith = false)
        {
            this.dte.ExecuteCommand("File.OpenFile", filePath);
        }

        private void saveCommands_AfterExecute(string Guid, int ID, object CustomIn, object CustomOut)
        {
            switch ((uint)ID)
            {
                case (uint)Microsoft.VisualStudio.VSConstants.VSStd97CmdID.SaveProjectItem:
                        UpdateProjFile(dte.ActiveWindow.Document.Path + dte.ActiveWindow.Document.Name);
                        break;
                case (uint)Microsoft.VisualStudio.VSConstants.VSStd97CmdID.SaveSolution:
                    foreach (string tempProjFile in this.tempToProjFiles.Keys)
                    {
                        UpdateProjFile(tempProjFile);
                    }
                    break;
                default:
                    return;
            }
        }


        private void UpdateProjFile(string tempFilePath)
        {
            string contents = ReadFile(tempFilePath);
            SetFileContents(this.tempToProjFiles[tempFilePath], contents);
        }

        private static void SetFileContents(string filePath, string content)
        {
            File.WriteAllText(filePath, content);
        }

        private static string ReadFile(string filePath)
        {
            return File.ReadAllText(filePath);
        }
    }
}
