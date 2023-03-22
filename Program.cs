using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
namespace WatchEmail
{
    class Program
    {
        static void Main(string[] args)
        {
            Program P = new();
            MAPIFolder? watchFolder = null;
            MAPIFolder inbox = P.GetOutlookInstance().GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            foreach (MAPIFolder subFolder in inbox.Folders)
            {
                if (subFolder.Name == "My Tickets") watchFolder = subFolder;
                if (watchFolder != null ) break;
            }
            
            if (watchFolder != null) watchFolder.Items.ItemAdd += new ItemsEvents_ItemAddEventHandler(P.Items_ItemAdd);

            // Keep the console application running
            Console.ReadLine();
        }
        public void Items_ItemAdd(object Item)
        {
            string filter = "will produce a court order for its release.";
            MailItem mail = (MailItem)Item;

            if (mail != null)
            {
                if (mail.Body.ToLower().Contains(filter))
                {
                    var proc = new Process
                    {
                        StartInfo = new ProcessStartInfo
                        {
                            FileName = @"C:\Source\GetNZip\GetNZip\bin\Debug\p.exe",
                            Arguments = String.Empty,
                            WorkingDirectory = @"C:\Source\GetNZip\GetNZip\bin\Debug\",
                            UseShellExecute = false,
                            RedirectStandardOutput = true,
                            CreateNoWindow = false,
                        }
                    };
                    proc.Start();
                    string stdoutx = proc.StandardOutput.ReadToEnd();
                    proc.WaitForExit();

                    Console.WriteLine("Stdout : {0}", stdoutx);
                }
            }
        }
        protected NameSpace GetOutlookInstance()
        {
            Application application;
            NameSpace nameSpace;
            application = new Application();
            nameSpace = application.GetNamespace("MAPI");
            nameSpace.Logon("Outlook", "", Missing.Value, Missing.Value);
            return nameSpace;
        }
    }
}