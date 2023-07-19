using System.Diagnostics;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
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
            MAPIFolder delFolder = pg.GetOutlookInstance().GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems);

            foreach (MAPIFolder subFolder in inbox.Folders)
            {
                if (subFolder.Name == "My Tickets") watchFolder = subFolder;
                if (watchFolder != null) break;
            }
            if (watchFolder != null) watchFolder.Items.ItemAdd += new ItemsEvents_ItemAddEventHandler(P.Items_ItemAdd);
            if (delFolder != null) delFolder.Items.ItemAdd += new ItemsEvents_ItemAddEventHandler(P.Pdf_ItemAdd);

            FileSystemWatcher bWatcher = new()
            {
                Path = (ParseNGet.Program.RemoteSave ? ParseNGet.Program.Constants.bobFolder.Replace("C:", "Y:") : ParseNGet.Program.Constants.bobFolder),
                NotifyFilter = NotifyFilters.LastWrite,
                Filter = "*-Msg.txt"
            };
            bWatcher.Changed += new FileSystemEventHandler(P.OnbCreated);
            bWatcher.EnableRaisingEvents = true;

            FileSystemWatcher jWatcher = new()
            {
                Path = (ParseNGet.Program.RemoteSave ? ParseNGet.Program.Constants.jingerFolder.Replace("C:", "Y:") : ParseNGet.Program.Constants.jingerFolder),
                NotifyFilter = NotifyFilters.LastWrite,
                Filter = "*-Msg.txt"
            };
            jWatcher.Changed += new FileSystemEventHandler(P.OnjCreated);
            jWatcher.EnableRaisingEvents = true;

            if (args.Length == 1 && args[0] == "o")
            {
                P.MakeAppointment();
                return;
            }

            // Keep the console application running
            Console.ReadLine();
        }
        protected void OnbCreated(object source, FileSystemEventArgs e)
        {
            ParseNGet.Program pg = new();
            pg.UpdateNoDataOrAttach(System.IO.Path.GetDirectoryName(e.FullPath));
        }
        protected void OnjCreated(object source, FileSystemEventArgs e)
        {
            ParseNGet.Program pg = new();
            pg.UpdateJingerDrafts(System.IO.Path.GetDirectoryName(e.FullPath));
        }
        protected void Items_ItemAdd(object Item)
        {
            MailItem mail = (MailItem)Item;

            string fileName = @"C:\Source\GetNZip\GetNZip\bin\Debug\p.exe";
            string arg = "n";

            if (mail != null)
            {
                string mailBody = mail.Body;

                bool forScoff = ParseNGet.Program.Constants.keywordsScoff.Any(s => mail.SenderEmailAddress.Contains(s));
                bool forBob = ParseNGet.Program.Constants.keywordsBob.Any(s => mailBody.Contains(s));
                bool forBob2 = ParseNGet.Program.Constants.keywordsBob2.Any(s => mailBody.ToLower().Contains(s));
                bool forJinger = ParseNGet.Program.Constants.keywordsJinger.Any(s => mailBody.Contains(s));
                bool findBal = ParseNGet.Program.Constants.keywordsBal.Any(s => mailBody.ToLower().Contains(s));
                bool listVeh = ParseNGet.Program.Constants.keywordsListVeh.Any(s => mailBody.ToLower().Contains(s));
                bool locateChk = ParseNGet.Program.Constants.keywordsChk.Any(s => mailBody.ToLower().Contains(s));
                bool forDon = ParseNGet.Program.Constants.keywordsDon.Any(s => mailBody.ToLower().Contains(s));
                bool forSeema = ParseNGet.Program.Constants.keywordsSeema.Any(s => mailBody.ToLower().Contains(s));
                bool provideVIN = ParseNGet.Program.Constants.keywordsVIN.Any(s => mailBody.Contains(s));
                bool provideLS = ParseNGet.Program.Constants.keywordsLS.Any(s => mailBody.Contains(s));
                bool noArg = forBob || findBal || listVeh || locateChk || forDon || forScoff;

                if (provideVIN) arg = "v";
                else if (provideLS) arg = "ls";
                else if (noArg) arg = string.Empty;

                if (forJinger || forSeema || forBob2)
                {
                    fileName = @"C:\Source\GetNZip\GetNZip\bin\Debug\z.exe";
                    arg = string.Empty;
                    if (forSeema) arg = "s";
                    else if (forBob2) arg = "b";
                }

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
                //proc.BeginOutputReadLine();
                string stdoutx = proc.StandardOutput.ReadToEnd();
                Console.WriteLine("Stdout : {0}", stdoutx);
                proc.WaitForExit();
                proc.Close();
            }
        }
        protected void Pdf_ItemAdd(object Item)
        {
            ParseNGet.Program pg = new();
            MAPIFolder inbox = pg.GetOutlookInstance().GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            MAPIFolder? pdfFolder = null;

            foreach (MAPIFolder subFolder in inbox.Folders)
            {
                if (subFolder.Name == "IT-EA") pdfFolder = subFolder;
                if (pdfFolder != null) break;
            }

            MailItem mail = (MailItem)Item;
            if(mail.Attachments.Count >0)
                foreach (Attachment attachment in mail.Attachments)
                {
                    if (attachment.FileName.ToLower().EndsWith(".pdf"))
                    {
                        string attachmentPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), attachment.FileName);
                        attachment.SaveAsFile(attachmentPath);
                        PdfReader reader = new(attachmentPath);

                        for (int i = 1; i <= reader.NumberOfPages; i++)
                        {
                            string[] lines = PdfTextExtractor.GetTextFromPage(reader, i).Split('\n');

                            foreach (string line in lines)
                            {
                                if (line.Contains(" Information Technology "))
                                {
                                    Console.WriteLine(line);
                                }
                            }
                        }
                        Console.WriteLine('\n' +"no one from IT quits.");
                        reader.Close();
                        File.Delete(attachmentPath);
                    }
                }
            else
                Console.WriteLine('\n' + "no attachment detected.");
            mail.Move(pdfFolder);
        }
        protected void MakeAppointment()
        {
            Application outlook = new Application();
            AppointmentItem appointment = outlook.CreateItem(OlItemType.olAppointmentItem);

            appointment.Subject = "Huaping Yu OOO";
            appointment.Location = "OOO";
            appointment.Start = new DateTime(2023, 04, 22, 08, 30, 0);
            appointment.End = new DateTime(2023, 04, 22, 17, 30, 0);

            Recipient optionalAttendee = appointment.Recipients.Add("IT-TollingAppSupport@ntta.org");
            optionalAttendee.Type = (int)OlMeetingRecipientType.olOptional;

            appointment.Save();

            MailItem mail = appointment.ForwardAsVcal();
            mail.Recipients.Add("HYu@ntta.org");

            //mail.Recipients.Add("IT-TollingAppSupport@ntta.org");

            mail.DeferredDeliveryTime = appointment.Start.Subtract(TimeSpan.FromHours(8));
            mail.Send();
        }
    }
}