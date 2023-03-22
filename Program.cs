using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
namespace WatchEmail
{
    class Program
    {
        public static class Constants
        {
            public static readonly string[] keywordsBal = { " outstanding balance", " outstanding toll", " zc balance" };
            public static readonly string[] keywordsListVeh = { " list of vehicles ", " list of transponders " };
            public static readonly string[] keywordsChk = { " unable to locate check ", "not be located ", " cannot locate ", " ck #", " check #", " check#" };
            public static readonly string[] keywordsDon = { " anil has code for this ", " need hv and vrb data ", " need data pulled soon after " };
            public static readonly string[] keywordsSeema = { "img pull", "img request", "image pull", "image request", "pull images" };
            public static readonly string[] keywordsVIN = { " Please provide VIN", " for the attached plates" };
            public static readonly string[] keywordsBob = { "Email: rdigman@ntta.org" };
            public static readonly string[] keywordsJinger = { "Email: Jelmore@ntta.org" };
            public static readonly string keywordsScoff = "VTR_SCOFFLAW@txdmv.gov";
        }
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
                bool findBal = Constants.keywordsBal.Any(s => mail.Body.ToLower().Contains(s));
                bool listVeh = Constants.keywordsListVeh.Any(s => mail.Body.ToLower().Contains(s));
                bool locateChk = Constants.keywordsChk.Any(s => mail.Body.ToLower().Contains(s));
                bool forDon = Constants.keywordsDon.Any(s => mail.Body.ToLower().Contains(s));
                bool forSeema = Constants.keywordsSeema.Any(s => mail.Body.ToLower().Contains(s));
                bool provideVIN = Constants.keywordsVIN.Any(s => mail.Body.Contains(s));
                bool fromScoff = mail.SenderEmailAddress == Constants.keywordsScoff;
                bool fromBob = Constants.keywordsBob.Any(s => mail.Body.Contains(s));
                bool fromJinger = Constants.keywordsJinger.Any(s => mail.Body.Contains(s));

                if (fromJinger || forSeema)
                {
                    fileName = @"C:\Source\GetNZip\GetNZip\bin\Debug\z.exe";
                    if (forSeema) arg = "s";
                }

                if (provideVIN) arg = "v";

                if (fromBob || findBal || listVeh || locateChk || forDon) arg = string.Empty;

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