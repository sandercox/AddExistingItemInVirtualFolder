//------------------------------------------------------------------------------
// <copyright file="AddExistingItemPackage.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;

namespace AddExistingItemInVirtualFolder
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(AddExistingItemPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad("ADFC4E64-0397-11D1-9F4E-00A0C911004F")]
    public sealed class AddExistingItemPackage : Package
    {
        /// <summary>
        /// AddExistingItemPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "d51c3a56-13ce-413f-8f37-be226214fe45";

        private const string AddExistingItemGUID = "{5EFC7975-14BC-11CF-9B2B-00AA00573819}";
        private const int AddExistingItemCommandID = 244;

        private EnvDTE.CommandEvents commandEvents;

        /// <summary>
        /// Initializes a new instance of the <see cref="AddExistingItemPackage"/> class.
        /// </summary>
        public AddExistingItemPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();

            var DTE = GetService(typeof(EnvDTE.DTE)) as EnvDTE.DTE;
            commandEvents = DTE.Events.CommandEvents;
            commandEvents.BeforeExecute += delegate (string Guid, int ID, object CustomIn, object CustomOut, ref bool CancelDefault)
            {
                if (ID == AddExistingItemCommandID)
                {
                    if (Guid == AddExistingItemGUID)
                    {
                        // try to get a project item from where this command was executed
                        var projectItem = GetSelectedItem() as EnvDTE.ProjectItem;

                        if (projectItem != null && projectItem.Kind == EnvDTE.Constants.vsProjectItemKindVirtualFolder)
                        {
                            Debug.WriteLine($"Virtual folder item: {projectItem.Name}");

                            var folder = GetPhysicalFolderForVirtualFolder(projectItem);                            
                            if (folder != null) // we were able to determine a folder
                            {
                                // prepare open file dialog for any file
                                var dlg = new System.Windows.Forms.OpenFileDialog();
                                dlg.CheckFileExists = true;
                                dlg.CheckPathExists = true;
                                dlg.InitialDirectory = folder;
                                dlg.Multiselect = true;
                                dlg.Filter = "Any File|*.*";
                                dlg.Title = "Add Existing Item...";
                                
                                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                                {
                                    // if user selected one or more files add them to the project item and cancel the default action.
                                    Debug.WriteLine($"ADD EXISTING: {dlg.FileName}");
                                    foreach(var file in dlg.FileNames)
                                    {
                                        projectItem.ProjectItems.AddFromFile(file);
                                    }
                                    CancelDefault = true;
                                }
                            }
                        }
                    }
                }
            };
        }

        public static object GetSelectedItem()
        {
            IntPtr hierarchyPointer, selectionContainerPointer;
            object selectedObject = null;
            IVsMultiItemSelect multiItemSelect;
            uint itemId;

            var monitorSelection = (IVsMonitorSelection)Package.GetGlobalService(typeof(SVsShellMonitorSelection));

            try
            {
                monitorSelection.GetCurrentSelection(out hierarchyPointer,
                                                 out itemId,
                                                 out multiItemSelect,
                                                 out selectionContainerPointer);

                IVsHierarchy selectedHierarchy = Marshal.GetTypedObjectForIUnknown(
                                                     hierarchyPointer,
                                                     typeof(IVsHierarchy)) as IVsHierarchy;

                if (selectedHierarchy != null)
                {
                    ErrorHandler.ThrowOnFailure(selectedHierarchy.GetProperty(itemId, (int)__VSHPROPID.VSHPROPID_ExtObject, out selectedObject));
                }

                Marshal.Release(hierarchyPointer);
                Marshal.Release(selectionContainerPointer);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.Write(ex);
            }

            return selectedObject;
        }
        
        private static string GetPhysicalFolderForVirtualFolder(EnvDTE.ProjectItem item, bool exploreSubFolders = true)
        {
            if (item == null)
                return null;

            var items = item.ProjectItems;

            // this only searches if there is already a file in this directory - if not it will just fail to determine a path
            foreach (EnvDTE.ProjectItem it in items)
            {
                if (System.IO.File.Exists(it.FileNames[1]))
                {
                    return System.IO.Path.GetDirectoryName(it.FileNames[1]);
                }
            }

            if (exploreSubFolders)
            {
                // try to find sub virtual folder that have files in them
                foreach (EnvDTE.ProjectItem it in items)
                {
                    if (it.Kind == EnvDTE.Constants.vsProjectItemKindVirtualFolder)
                    {
                        var folder = GetPhysicalFolderForVirtualFolder(it);

                        if (folder != null)
                        {
                            var parent = System.IO.Directory.GetParent(folder);
                            if (parent.Name == item.Name)
                            {
                                return parent.FullName;
                            }
                        }
                    }
                }
            }

            // still no path try to search upward in the tree
            var folderParent = item.Collection?.Parent as EnvDTE.ProjectItem;
            if (folderParent != null && folderParent.Kind == EnvDTE.Constants.vsProjectItemKindVirtualFolder)
            {
                var folder = GetPhysicalFolderForVirtualFolder(folderParent, false);

                if (folder != null)
                {
                    var newFolder = System.IO.Path.Combine(folder, item.Name);
                    if (System.IO.Directory.Exists(newFolder))
                        return newFolder;
                }
            }

            return null;
        }
        

        #endregion
    }
}
