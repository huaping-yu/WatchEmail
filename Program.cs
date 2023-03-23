﻿using System.Diagnostics;
using System.Reflection;
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
        protected void Items_ItemAdd(object Item)
        {
            MailItem mail = (MailItem)Item;

            string arg = "n";
            string fileName = @"C:\Source\GetNZip\GetNZip\bin\Debug\p.exe";

            if (mail != null)
            {
                bool findBal = ParseNGet.Program.Constants.keywordsBal.Any(s => mail.Body.ToLower().Contains(s));
                bool listVeh = ParseNGet.Program.Constants.keywordsListVeh.Any(s => mail.Body.ToLower().Contains(s));
                bool locateChk = ParseNGet.Program.Constants.keywordsChk.Any(s => mail.Body.ToLower().Contains(s));
                bool forDon = ParseNGet.Program.Constants.keywordsDon.Any(s => mail.Body.ToLower().Contains(s));
                bool forSeema = ParseNGet.Program.Constants.keywordsSeema.Any(s => mail.Body.ToLower().Contains(s));
                bool provideVIN = ParseNGet.Program.Constants.keywordsVIN.Any(s => mail.Body.Contains(s));
                bool fromScoff = ParseNGet.Program.Constants.keywordsScoff.Any(s=> mail.SenderEmailAddress.Contains(s));
                bool fromBob = ParseNGet.Program.Constants.keywordsBob.Any(s => mail.Body.Contains(s));
                bool fromBob2 = ParseNGet.Program.Constants.keywordsBob2.Any(s => mail.Body.Contains(s));
                bool fromJinger = ParseNGet.Program.Constants.keywordsJinger.Any(s => mail.Body.Contains(s));

                if (fromJinger || forSeema || fromBob2)
                {
                    fileName = @"C:\Source\GetNZip\GetNZip\bin\Debug\z.exe";
                    if (forSeema) arg = "s";
                    if (fromBob2) arg = "b";
                }

                if (provideVIN) arg = "v";

                if (fromBob || findBal || listVeh || locateChk || forDon || fromScoff) arg = string.Empty;

                var proc = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = fileName,
                        Arguments = arg,
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