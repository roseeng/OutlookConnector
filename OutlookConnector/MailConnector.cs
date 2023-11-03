using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookConnector
{
    public class MailConnector
    {
        private static Outlook.Application outlookApp = null;

        /// <summary>
        /// Öppna en sparad msg-fil som mall. Uppdatera mottagare och lägg till en bilaga (förslagsvis Inbjudan)
        /// Poppar sedan upp mailet så att man kan granska och trycka på Skicka.
        /// </summary>
        /// <param name="msgfile">Fullständig sökväg till msg-fil</param>
        /// <param name="recipient">Epostadress till mottagare</param>
        /// <param name="attachmentfile">Fullständig sökväg till bilaga</param>
        /// <returns>En sträng med ett felmeddelande eller tom sträng om allt ok</returns>
        public static string OpenEmail(string msgfile, string recipient, string attachmentfile)
        {
            if (outlookApp == null)
            {
                try
                {
                    outlookApp = new Outlook.Application();
                }
                catch
                {
                    return "Hittar inte Outlook på denna datorn.";
                }
            }

            if (!File.Exists(msgfile))
            {
                return "Hittar inte mallfilen " + msgfile + ". Se till att den finns i samma mapp som PAKS.";
            }
            if (!File.Exists(attachmentfile))
            {
                return "Hittar inte bilagan '" + attachmentfile + "'";
            }

            var item = outlookApp.Session.OpenSharedItem(msgfile);
            if (!string.IsNullOrEmpty(recipient)) 
                item.Recipients.Add(recipient);
            item.Attachments.Add(attachmentfile, 1, 1, Path.GetFileName(attachmentfile));
            item.Display(false);

            return "";
        }
    }
}
