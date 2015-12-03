using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace rdomunozcom.EditProj
{
    /// <summary>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidEditProjPkgString)]
    public sealed class EditProjPackage : Package
    {
        private CommandEvents saveFileCommand, saveAllCommand, exitCommand;
        private DocumentEvents documentEvents;
        private DTE dte;
        private IDictionary<string, string> tempToProjFiles =
            new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);

        // Calling the QueryEditFiles method within the body pops up a query dialog. While the dialog is waiting for the user, the shell automatically can call the method checking for dirtiness and that will call CanEditFile again. To avoid recursion we use the _GettingCheckOutStatus flag as a guard
        private bool gettingCheckoutStatus = false;

        public EditProjPackage()
        {
        }

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
                var menuItem = new MenuCommand(MenuItemCallback, menuCommandID);

                this.dte = this.GetService(typeof(DTE)) as DTE;
                // need to keep a strong reference to CommandEvents so that it's not GC'ed
                this.saveFileCommand = this.dte.Events.CommandEvents[VSConstants.CMDSETID.StandardCommandSet97_string, (int)VSConstants.VSStd97CmdID.SaveProjectItem];
                this.saveAllCommand = this.dte.Events.CommandEvents[VSConstants.CMDSETID.StandardCommandSet97_string, (int)VSConstants.VSStd97CmdID.SaveSolution];
                this.exitCommand = this.dte.Events.CommandEvents[VSConstants.CMDSETID.StandardCommandSet97_string, (int)VSConstants.VSStd97CmdID.Exit];

                this.saveFileCommand.AfterExecute += saveCommands_AfterExecute;
                
                documentEvents = this.dte.Events.DocumentEvents;
                documentEvents.DocumentClosing += documentClosing;

                this.saveAllCommand.AfterExecute += saveCommands_AfterExecute;
                this.exitCommand.BeforeExecute += exitCommand_BeforeExecute;
                mcs.AddCommand(menuItem);
            }
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            try
            {
                string projFilePath = GetPathOfSelectedItem();
                string tempProjFilePath;

                if (TempFileExists(projFilePath, out tempProjFilePath))
                {
                    tempProjFilePath = GetTempFileFor(projFilePath);
                }
                else
                {
                    //can delete this if?
                    if (tempProjFilePath != null)
                    {
                        this.tempToProjFiles.Remove(tempProjFilePath);
                    }

                    tempProjFilePath = GetNewTempFilePath();
                    this.tempToProjFiles[tempProjFilePath] = projFilePath;
                }

                CreateTempFileWithContentsOf(projFilePath, tempProjFilePath);
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

        private static string GetNewTempFilePath()
        {
            string tempDir = Path.GetTempPath();
            string tempProjFile = Guid.NewGuid().ToString() + ".xml";
            string tempProjFilePath = Path.Combine(tempDir, tempProjFile);
            return tempProjFilePath;
        }

        private static void CreateTempFileWithContentsOf(string sourcePath, string destPath)
        {
            string projectData = File.ReadAllText(sourcePath);
            File.WriteAllText(destPath, projectData);
        }

        private bool TempFileExists(string projFilePath, out string tempProjFile)
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
                        UpdateProjFile(dte.ActiveWindow.Document.FullName);
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

        private void documentClosing(Document document)
        {
            if (this.tempToProjFiles.ContainsKey(document.FullName))
            {
                this.tempToProjFiles.Remove(document.FullName);
                File.Delete(document.FullName);
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

        private void NotifyForSave(string filePath)
        {
            int hr;
            IVsQueryEditQuerySave2 queryEditQuerySave = (IVsQueryEditQuerySave2)GetService(typeof(SVsQueryEditQuerySave));
            uint result;
            hr = queryEditQuerySave.QuerySaveFile(filePath, 0, null, out result);
        }

        private static void SetFileContents(string filePath, string content)
        {
            File.WriteAllText(filePath, content);
        }

        private static string ReadFile(string filePath)
        {
            return File.ReadAllText(filePath);
        }

        private bool CanEditFile(string filePath)
        {
            if (gettingCheckoutStatus) return false;

            try
            {
                gettingCheckoutStatus = true;

                IVsQueryEditQuerySave2 queryEditQuerySave =
                  (IVsQueryEditQuerySave2)GetService(typeof(SVsQueryEditQuerySave));

                string[] documents = { filePath };
                uint result;
                uint outFlags;

                int hr = queryEditQuerySave.QueryEditFiles(
                  0,
                  documents.Length,
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

        private void NotifyDocChanged(string filePath)
        {
            IVsFileChangeEx fileChangeEx =
              (IVsFileChangeEx)GetService(typeof(SVsFileChangeEx));
            int hr;
            hr = fileChangeEx.SyncFile(filePath);
        }
    }
}
