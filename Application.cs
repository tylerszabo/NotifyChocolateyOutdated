// Copyright (C) 2018 Tyler Szabo
//
// This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.
//
// This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using Microsoft.Win32;
using Windows.Data.Xml.Dom;
using Windows.UI.Notifications;

using static Vanara.PInvoke.Ole32;
using static Vanara.PInvoke.PropSys;
using static Vanara.PInvoke.Shell32;

using chocolatey;
using chocolatey.infrastructure.app;
using chocolatey.infrastructure.app.services;
using chocolatey.infrastructure.results;

namespace NotifyChocolateyOutdated
{
    class Application
    {
        private const string APP_NAME = "NotifyChocolateyOutdated";
        private const string APP_ID = APP_NAME + "_{b0ea947e-db9d-4322-961f-06653ffa2c58}";
        private const string MUTEX_NAME = APP_NAME + "_{6350937f-f94c-46a6-8ead-c10256edc00e}";

        private static readonly string EXE_PATH = System.Reflection.Assembly.GetEntryAssembly().Location;
        private static readonly string LINKFILE_PATH = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Programs), APP_NAME + ".lnk");
        private const string SHORTCUT_ARGUMENTS = "";

        private const string NOTIFICATION_SETTINGS_SUBKEY = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings";

        private static bool SetShortcut()
        {
            bool writeShortcut = false;

            PROPERTYKEY appModelIDKey = PROPERTYKEY.System.AppUserModel.ID;

            IShellLinkW shortcut = (IShellLinkW)new CShellLinkW();
            IPersistFile persistFile = (IPersistFile)shortcut;
            IPropertyStore propertyStore = (IPropertyStore)shortcut;

            try
            {
                persistFile.Load(LINKFILE_PATH, (int)Vanara.PInvoke.STGM.STGM_READ);
            }
            catch (System.IO.FileNotFoundException)
            {
                writeShortcut = true;

                // This is possibly first invocation so set registry in order to persist in Action Center
                ConfigureRegistry();
            }

            if (!writeShortcut)
            {
                StringBuilder curPath = new StringBuilder(Vanara.PInvoke.Kernel32.MAX_PATH);
                shortcut.GetPath(curPath, curPath.Capacity, null, SLGP.SLGP_RAWPATH);

                if (!EXE_PATH.Equals(curPath.ToString(), StringComparison.OrdinalIgnoreCase))
                {
                    writeShortcut = true;
                }
            }

            if (!writeShortcut)
            {
                StringBuilder curArgs = new StringBuilder(Vanara.PInvoke.ComCtl32.INFOTIPSIZE);
                shortcut.GetArguments(curArgs, curArgs.Capacity);

                if (!SHORTCUT_ARGUMENTS.Equals(curArgs.ToString(), StringComparison.Ordinal))
                {
                    writeShortcut = true;
                }
            }

            if (!writeShortcut)
            {
                using (PROPVARIANT pv = new PROPVARIANT())
                {
                    propertyStore.GetValue(ref appModelIDKey, pv);
                    if (!APP_ID.Equals(pv.pwszVal, StringComparison.Ordinal))
                    {
                        writeShortcut = true;
                    }
                }
            }

            if (writeShortcut)
            {
                shortcut.SetPath(EXE_PATH);
                shortcut.SetArguments(string.Empty);

                using (PROPVARIANT pv = new PROPVARIANT(APP_ID))
                {
                    propertyStore.SetValue(ref appModelIDKey, pv);
                    propertyStore.Commit();
                }

                persistFile.Save(LINKFILE_PATH, true);

                return true;
            }

            return false;
        }

        private static void ConfigureRegistry()
        {
            using (RegistryKey systemNotificationSettings = Registry.LocalMachine.OpenSubKey(NOTIFICATION_SETTINGS_SUBKEY, RegistryKeyPermissionCheck.ReadSubTree))
            {
                HashSet<string> subKeys = new HashSet<string>(systemNotificationSettings.GetSubKeyNames());
                if (subKeys.Contains(APP_ID))
                {
                    // Configuration already exists system-wide, don't overwrite
                    return;
                }
            }

            using (RegistryKey notificationSettings = Registry.CurrentUser.OpenSubKey(NOTIFICATION_SETTINGS_SUBKEY, RegistryKeyPermissionCheck.ReadWriteSubTree))
            {
                HashSet<string> subKeys = new HashSet<string>(notificationSettings.GetSubKeyNames());
                if (subKeys.Contains(APP_ID))
                {
                    // Configuration already exists for user, don't overwrite
                    return;
                }
                using (RegistryKey appSettings = notificationSettings.CreateSubKey(APP_ID))
                {
                    appSettings.SetValue("ShowInActionCenter", 1, RegistryValueKind.DWord);
                }
            }
        }

        private const string TOAST_TITLE = "Outdated Chocolatey Packages";
        private const string TOAST_BODY_SINGLE = "{0} outdated package";
        private const string TOAST_BODY_MULTIPLE = "{0} outdated packages";

        private static void ShowToast(int packageCount)
        {
            ToastNotificationManager.History.Clear(APP_ID);

            if (packageCount > 0)
            {
                XmlDocument toastXml = ToastNotificationManager.GetTemplateContent(ToastTemplateType.ToastText02);

                XmlNodeList textNodes = toastXml.GetElementsByTagName("text");
                textNodes[0].AppendChild(toastXml.CreateTextNode(TOAST_TITLE));
                textNodes[1].AppendChild(toastXml.CreateTextNode(string.Format(packageCount == 1 ? TOAST_BODY_SINGLE : TOAST_BODY_MULTIPLE, packageCount)));

                ToastNotification toast = new ToastNotification(toastXml);
                toast.ExpirationTime = DateTimeOffset.Now.AddDays(2);

                ToastNotificationManager.CreateToastNotifier(APP_ID).Show(toast);
            }
        }

        // This can take a while with many packages
        private static int GetOutdatedCount()
        {
            GetChocolatey chocolatey = Lets.GetChocolatey();
            chocolatey.SetCustomLogging(new chocolatey.infrastructure.logging.NullLog());
            chocolatey.Set(c => { c.CommandName = "outdated"; c.PackageNames = ApplicationParameters.AllPackages; });
            INugetService nuget = chocolatey.Container().GetInstance<INugetService>();

            ConcurrentDictionary<string, PackageResult> upgradeResults = nuget.upgrade_noop(chocolatey.GetConfiguration(), null);

            return upgradeResults.Values.Count((v) => { return !v.Inconclusive; });
        }

        [STAThread]
        public static void Main()
        {
            using (Mutex mutext = new Mutex(false, MUTEX_NAME))
            {
                if (mutext.WaitOne(0, true))
                {
                    // Allow only one instance at a time

                    if (SetShortcut())
                    {
                        // Have to restart if Shortcut is updated
                        System.Windows.Forms.Application.Restart();
                        return;
                    }

                    int outdatedPackageCount = GetOutdatedCount();

                    ShowToast(outdatedPackageCount);
                }
            }
        }
    }
}