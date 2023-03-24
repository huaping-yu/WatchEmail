using System.Diagnostics;
using Microsoft.Office.Interop.Outlook;
namespace WatchEmail
{
    class Program
    {
        static void Main(string[] args)
        {
            Program P = new();
            ParseNGet.Program pg = new();
            MAPIFolder? watchFolder = null;
            MAPIFolder inbox = pg.GetOutlookInstance().GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            foreach (MAPIFolder subFolder in inbox.Folders)
            {
                if (subFolder.Name == "My Tickets") watchFolder = subFolder;
                if (watchFolder != null ) break;
            }
            if (watchFolder != null) watchFolder.Items.ItemAdd += new ItemsEvents_ItemAddEventHandler(P.Items_ItemAdd);

            FileSystemWatcher bWatcher = new()
            {
                Path = ParseNGet.Program.Constants.bobFolder,
                NotifyFilter = NotifyFilters.LastWrite,
                Filter = "*-Msg.txt"
            };
            bWatcher.Changed += new FileSystemEventHandler(P.OnbCreated);
            bWatcher.EnableRaisingEvents = true;

            FileSystemWatcher jWatcher = new()
            {
                Path = ParseNGet.Program.Constants.jingerFolder,
                NotifyFilter = NotifyFilters.LastWrite,
                Filter = "*-Msg.txt"
            };
            jWatcher.Changed += new FileSystemEventHandler(P.OnjCreated);
            jWatcher.EnableRaisingEvents = true;

            // Keep the console application running
            Console.ReadLine();
        }
        protected void OnbCreated(object source, FileSystemEventArgs e)
        {
            ParseNGet.Program pg = new();
            pg.UpdateNoDataAttach(ParseNGet.Program.Constants.bobFolder);
        }
        protected void OnjCreated(object source, FileSystemEventArgs e)
        {
            ParseNGet.Program pg = new();
            pg.UpdateNoDataAttach(ParseNGet.Program.Constants.jingerFolder);
        }
        protected void Items_ItemAdd(object Item)
        {
            MailItem mail = (MailItem)Item;
           
            string fileName = @"C:\Source\GetNZip\GetNZip\bin\Debug\p.exe";
            string arg = "n";

            if (mail != null)
            {
                string mailBody = mail.Body;
                bool findBal = ParseNGet.Program.Constants.keywordsBal.Any(s => mailBody.ToLower().Contains(s));
                bool listVeh = ParseNGet.Program.Constants.keywordsListVeh.Any(s => mailBody.ToLower().Contains(s));
                bool locateChk = ParseNGet.Program.Constants.keywordsChk.Any(s => mailBody.ToLower().Contains(s));
                bool forDon = ParseNGet.Program.Constants.keywordsDon.Any(s => mailBody.ToLower().Contains(s));
                bool forSeema = ParseNGet.Program.Constants.keywordsSeema.Any(s => mailBody.ToLower().Contains(s));
                bool provideVIN = ParseNGet.Program.Constants.keywordsVIN.Any(s => mailBody.Contains(s));
                bool fromScoff = ParseNGet.Program.Constants.keywordsScoff.Any(s=> mail.SenderEmailAddress.Contains(s));
                bool fromBob = ParseNGet.Program.Constants.keywordsBob.Any(s => mailBody.Contains(s));
                bool fromBob2 = ParseNGet.Program.Constants.keywordsBob2.Any(s => mailBody.Contains(s));
                bool fromJinger = ParseNGet.Program.Constants.keywordsJinger.Any(s => mailBody.Contains(s));

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
    }
}