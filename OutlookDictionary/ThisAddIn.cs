using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;

namespace OutlookDictionary
{


    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
            new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);


        }
        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID != null)
                {
                    mailItem.HTMLBody = mailItem.HTMLBody.Replace("ggfsdt", "<a href='Go Go Forward So Dis Toward'>ggfsdt</a>");
                    MessageBox.Show("asd");
                }

            }



        }
        /*private string GetDataFile()
        {
            string userSettingsPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

            // build company folder full path
            string companyFolder = Path.Combine(userSettingsPath, "OutlookDictionary");

            if (!Directory.Exists(companyFolder))
            {
                Directory.CreateDirectory(companyFolder);

                string fullSettingsPath = Path.Combine(companyFolder, "abbrev-list.csv");

                return fullSettingsPath;
            }


            userSettingsPath = Path.Combine(companyFolder, "abbrev-list.csv");
            return userSettingsPath;
        }*/

        
    }
}