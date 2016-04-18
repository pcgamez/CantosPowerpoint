using System;
using System.Collections;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using IWshRuntimeLibrary;
//using Shell32; // Need to import Shell32?

namespace CantosPowerpoint
{
    [RunInstaller(true)]
    public partial class Installer : System.Configuration.Install.Installer
    {

        #region Private Instance Variables

        private string _location = null;
        private string _name = null;
        private string _description = null;

        #endregion

        #region Private Properties

        private string QuickLaunchFolder
        {
            get
            {
                return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) +
                    "\\Microsoft\\Internet Explorer\\Quick Launch";
            }
        }


        private string ShortcutTarget
        {
            get
            {
                if (_location == null)
                    _location = Assembly.GetExecutingAssembly().Location;
                return _location;
            }
        }


        private string ShortcutName
        {
            get
            {
                if (_name == null)
                {
                    Assembly myAssembly = Assembly.GetExecutingAssembly();

                    try
                    {
                        object titleAttribute = myAssembly.GetCustomAttributes(typeof(AssemblyTitleAttribute), false)[0];
                        _name = ((AssemblyTitleAttribute)titleAttribute).Title;
                    }
                    catch { }

                    if ((_name == null) || (_name.Trim() == string.Empty))
                        _name = myAssembly.GetName().Name;
                }
                return _name;
            }
        }


        private string ShortcutDescription
        {
            get
            {
                if (_description == null)
                {
                    Assembly myAssembly = Assembly.GetExecutingAssembly();

                    try
                    {
                        object descriptionAttribute = myAssembly.GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false)[0];
                        _description = ((AssemblyDescriptionAttribute)descriptionAttribute).Description;
                    }
                    catch { }

                    if ((_description == null) || (_description.Trim() == string.Empty))
                        _description = "Launch " + ShortcutName;
                }
                return _description;
            }
        }

        #endregion

        #region Override Methods
        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);

            const string DESKTOP_SHORTCUT_PARAM = "DESKTOP_SHORTCUT";
            const string QUICKLAUNCH_SHORTCUT_PARAM = "QUICKLAUNCH_SHORTCUT";
            const string ALLUSERS_PARAM = "ALLUSERS";

            // The installer will pass the ALLUSERS, DESKTOP_SHORTCUT_PARAM and QUICKLAUNCH_SHORTCUT   
            // parameters. These have been set to the values of radio buttons and checkboxes from the
            // MSI user interface.
            // ALLUSERS is set according to whether the user chooses to install for all users (="1") 
            // or just for themselves (="").
            // If the user checked the checkbox to install one of the shortcuts, then the corresponding 
            // parameter value is "1".  If the user did not check the checkbox to install one of the 
            // desktop shortcut, then the corresponding parameter value is an empty string.

            // First make sure the parameters have been provided.
            if (!Context.Parameters.ContainsKey(ALLUSERS_PARAM))
                throw new Exception(string.Format("The {0} parameter has not been provided for the {1} class.", ALLUSERS_PARAM, this.GetType()));
            if (!Context.Parameters.ContainsKey(DESKTOP_SHORTCUT_PARAM))
                throw new Exception(string.Format("The {0} parameter has not been provided for the {1} class.", DESKTOP_SHORTCUT_PARAM, this.GetType()));
            if (!Context.Parameters.ContainsKey(QUICKLAUNCH_SHORTCUT_PARAM))
                throw new Exception(string.Format("The {0} parameter has not been provided for the {1} class.", QUICKLAUNCH_SHORTCUT_PARAM, this.GetType()));
            bool allusers = Context.Parameters[ALLUSERS_PARAM] != string.Empty;
            bool installDesktopShortcut = Context.Parameters[DESKTOP_SHORTCUT_PARAM] != string.Empty;
            bool installQuickLaunchShortcut = Context.Parameters[QUICKLAUNCH_SHORTCUT_PARAM] != string.Empty;

            if (installDesktopShortcut)
            {
                // If this is an All Users install then we need to install the desktop shortcut for 
                // all users.  .Net does not give us access to the All Users Desktop special folder,
                // but we can get this using the Windows Scripting Host.
                string desktopFolder = null;

                if (allusers)
                {
                    try
                    {
                        // This is in a Try block in case AllUsersDesktop is not supported
                        object allUsersDesktop = "AllUsersDesktop";
                        WshShell shell = new WshShell();
                        desktopFolder = shell.SpecialFolders.Item(ref allUsersDesktop).ToString();
                    }
                    catch { }
                }
                if (desktopFolder == null)
                    desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

                CreateShortcut(desktopFolder, ShortcutName, ShortcutTarget, ShortcutDescription);
            }

            if (installQuickLaunchShortcut)
            {
                CreateShortcut(QuickLaunchFolder, ShortcutName, ShortcutTarget, ShortcutDescription);
            }

        }

        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);

            DeleteShortcuts();
        }


        public override void Rollback(IDictionary savedState)
        {
            base.Rollback(savedState);

            DeleteShortcuts();
        }

        #endregion

        #region Private Helper Methods

/*        private static void PinUnpinTaskBar(string filePath, bool pin)
        {
            if (!System.IO.File.Exists(filePath)) throw new FileNotFoundException(filePath);

            // create the shell application object
            Shell shellApplication = new Shell();

            string path = Path.GetDirectoryName(filePath);
            string fileName = Path.GetFileName(filePath);

            Shell32.Folder directory = shellApplication.NameSpace(path);
            FolderItem link = directory.ParseName(fileName);

            FolderItemVerbs verbs = link.Verbs();
            for (int i = 0; i < verbs.Count; i++)
            {
                FolderItemVerb verb = verbs.Item(i);
                string verbName = verb.Name.Replace(@"&", string.Empty).ToLower();

                if ((pin && verbName.Equals("pin to taskbar")) || (!pin && verbName.Equals("unpin from taskbar")))
                {
                    verb.DoIt();
                }
            }

            shellApplication = null;
        }*/

        private void CreateShortcut(string folder, string name, string target, string description)
        {
            string shortcutFullName = Path.Combine(folder, name + ".lnk");

            try
            {
                WshShell shell = new WshShell();
                IWshShortcut link = (IWshShortcut)shell.CreateShortcut(shortcutFullName);
                link.TargetPath = target;
                link.Description = description;
                link.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("The shortcut \"{0}\" could not be created.\n\n{1}", shortcutFullName, ex.ToString()),
                    "Create Shortcut", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }


        private void DeleteShortcuts()
        {
            // Just try and delete all possible shortcuts that may have been
            // created during install

            try
            {
                // This is in a Try block in case AllUsersDesktop is not supported
                object allUsersDesktop = "AllUsersDesktop";
                WshShell shell = new WshShell();
                string desktopFolder = shell.SpecialFolders.Item(ref allUsersDesktop).ToString();
                DeleteShortcut(desktopFolder, ShortcutName);
            }
            catch { }

            DeleteShortcut(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), ShortcutName);

            DeleteShortcut(QuickLaunchFolder, ShortcutName);
        }


        private void DeleteShortcut(string folder, string name)
        {
            string shortcutFullName = Path.Combine(folder, name + ".lnk");
            FileInfo shortcut = new FileInfo(shortcutFullName);
            if (shortcut.Exists)
            {
                try
                {
                    shortcut.Delete();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("The shortcut \"{0}\" could not be deleted.\n\n{1}", shortcutFullName, ex.ToString()),
                        "Delete Shortcut", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }


        #endregion

    }
}
