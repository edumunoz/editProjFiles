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
using System.Linq;
using EnvDTE80;

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

        // Calling the QueryEditFiles method within the body pops up a query dialog. While the dialog is waiting for the user, the shell automatically can call the method checking for dirtiness and that will call CanEditFile again. To avoid recursion we use the _GettingCheckOutStatus flag as a guard
        private bool gettingCheckoutStatus = false;

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

                this.dte = this.GetService(typeof(DTE)) as DTE;
                // need to keep a strong reference to CommandEvents so that it's not GC'ed
                this.saveFileCommand = this.dte.Events.CommandEvents[VSConstants.CMDSETID.StandardCommandSet97_string, (int)VSConstants.VSStd97CmdID.SaveProjectItem];
                this.saveAllCommand = this.dte.Events.CommandEvents[VSConstants.CMDSETID.StandardCommandSet97_string, (int)VSConstants.VSStd97CmdID.SaveSolution];
                this.exitCommand = this.dte.Events.CommandEvents[VSConstants.CMDSETID.StandardCommandSet97_string, (int)VSConstants.VSStd97CmdID.Exit];

                this.saveFileCommand.AfterExecute += saveCommands_AfterExecute;
                this.saveAllCommand.AfterExecute += saveCommands_AfterExecute;
                this.exitCommand.BeforeExecute += exitCommand_BeforeExecute;
                mcs.AddCommand(menuItem);
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
            try
            {
                string projFilePath = GetPathOfSelectedItem();
                string tempProjFilePath;

                if (DocumentAlreadyOpen(projFilePath, out tempProjFilePath))
                {
                    tempProjFilePath = GetTempFileFor(projFilePath);
                }
                else
                {
                    if (tempProjFilePath != null)
                    {
                        this.tempToProjFiles.Remove(tempProjFilePath);
                    }

                    tempProjFilePath = CreateTempFileWithContentsOf(projFilePath);
                    this.tempToProjFiles[tempProjFilePath] = projFilePath;
                }

                OpenFileInEditor(tempProjFilePath);
            }

            catch (Exception ex)
            {
                Debug.WriteLine(string.Format("There was an exception: {0}", ex));
            }
        }

        private string GetTempFileFor(string projFilePath)
        {
            return this.tempToProjFiles.FirstOrDefault(kv => kv.Value == projFilePath).Key;
        }

        private string GetPathOfSelectedItem()
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
            return projFilePath;
        }

        private static string CreateTempFileWithContentsOf(string filePath)
        {
            string projectData = File.ReadAllText(filePath);
            string tempDir = Path.GetTempPath();
            string tempProjFile = Guid.NewGuid().ToString() + ".xml";
            string tempProjFilePath = Path.Combine(tempDir, tempProjFile);
            File.WriteAllText(tempProjFilePath, projectData);

            return tempProjFilePath;
        }

        private bool DocumentAlreadyOpen(string projFilePath, out string tempProjFile)
        {
            bool haveOpened = this.tempToProjFiles.Values.Contains(projFilePath);
            tempProjFile = null;
            if (haveOpened)
            {
                tempProjFile = GetTempFileFor(projFilePath);
                return File.Exists(tempProjFile);
            }

            return false;
        }


        private void OpenFileInEditor(string filePath)
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

        private void exitCommand_BeforeExecute(string Guid, int ID, object CustomIn, object CustomOut, ref bool CancelDefault)
        {
            foreach (Document doc in this.dte.Documents)
            {
                if (this.tempToProjFiles.ContainsKey(doc.FullName))
                {
                    string path = doc.FullName;
                    doc.Close();
                    File.Delete(path);
                }
            }
        }

        private void UpdateProjFile(string tempFilePath)
        {
            string contents = ReadFile(tempFilePath);
            string projFile = this.tempToProjFiles[tempFilePath];
            if (CanEditFile(this.tempToProjFiles[tempFilePath]))
            {
                NotifyForSave(projFile);
                SetFileContents(projFile, contents);
                NotifyDocChanged(projFile);
            }
        }

        private void NotifyForSave(string p)
        {
            int hr;
            IVsQueryEditQuerySave2 queryEditQuerySave = (IVsQueryEditQuerySave2)GetService(typeof(SVsQueryEditQuerySave));
            uint result;
            hr = queryEditQuerySave.QuerySaveFile(p, 0, null, out result);

        }

        private static void SetFileContents(string filePath, string content)
        {
            File.WriteAllText(filePath, content);
        }

        private static string ReadFile(string filePath)
        {
            return File.ReadAllText(filePath);
        }

        private bool CanEditFile(string p)
        {
            if (gettingCheckoutStatus) return false;

            try
            {
                gettingCheckoutStatus = true;

                IVsQueryEditQuerySave2 queryEditQuerySave =
                  (IVsQueryEditQuerySave2)GetService(typeof(SVsQueryEditQuerySave));

                string[] documents = { p };
                uint result;
                uint outFlags;

                int hr = queryEditQuerySave.QueryEditFiles(
                  0,
                  1,
                  documents,
                  null,
                  null,
                  out result,
                  out outFlags);
                if (ErrorHandler.Succeeded(hr) && (result ==
                  (uint)tagVSQueryEditResult.QER_EditOK))
                {
                    return true;
                }
            }
            finally
            {
                gettingCheckoutStatus = false;
            }
            return false;
        }


        private void NotifyDocChanged(string p)
        {
            IVsFileChangeEx fileChangeEx =
              (IVsFileChangeEx)GetService(typeof(SVsFileChangeEx));
            int hr;
            hr = fileChangeEx.SyncFile(p);
        }
    }
}
