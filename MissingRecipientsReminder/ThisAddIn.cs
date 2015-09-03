using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace MissingRecipientsReminder
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.ItemSend += Application_ItemSend;
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            Outlook.MailItem mailItem = item as Outlook.MailItem;
            if (mailItem == null)
            {
                return;
            }

            string body = mailItem.Body ?? string.Empty;

            string[] lines = body.Replace("\r\n", "\n").Split('\n');

            var missingRecipients = new List<string>();
            bool foundFirst = false;
            foreach (string line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                List<string> addedInBodyRecipients = ExtractRecipientsFromLine(line);
                if (addedInBodyRecipients != null && addedInBodyRecipients.Count > 0)
                {
                    missingRecipients.AddRange(addedInBodyRecipients);
                    foundFirst = true;
                }
                else
                {
                    if (foundFirst)
                    {
                        // No more recipient lines, so we stop before reading all the content
                        break;
                    }
                }
            }

            foreach (Outlook.Recipient recipient in mailItem.Recipients)
            {
                var alreadyIncludedRecipients = missingRecipients
                    .Where(r => this.RecipientContainsString(recipient, r))
                    .ToList();

                foreach (var alreadyIncludedRecipient in alreadyIncludedRecipients)
                {
                    missingRecipients.Remove(alreadyIncludedRecipient);
                }
            }

            if (missingRecipients.Count > 0)
            {
                var msg = new StringBuilder("You may have forgotten to add the following people:");
                foreach (var missingRecipient in missingRecipients)
                {
                    msg.AppendFormat("{0}   {1}", Environment.NewLine, missingRecipient);
                }
                msg.AppendLine();
                msg.AppendLine();
                msg.Append("Send anyway?");

                var result = MessageBox.Show(msg.ToString(), "Missing recipients", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (result != DialogResult.Yes)
                {
                    cancel = true;
                }
            }
        }

        private bool RecipientContainsString(Outlook.Recipient recipient, string name)
        {
            Outlook.AddressEntry addressEntry = recipient.AddressEntry;
            if (addressEntry == null)
            {
                return false;
            }

            if (addressEntry.Name.ContainsWholeWord(name, ignoreCase: true))
            {
                return true;
            }

            Outlook.ContactItem contact = addressEntry.GetContact();
            if (contact != null)
            {
                string email = contact.Email1Address;
                if (email != null && email.ContainsWholeWord(name, ignoreCase: true))
                {
                    return true;
                }

                email = contact.Email2Address;
                if (email != null && email.ContainsWholeWord(name, ignoreCase: true))
                {
                    return true;
                }

                email = contact.Email3Address;
                if (email != null && email.ContainsWholeWord(name, ignoreCase: true))
                {
                    return true;
                }

                string fullName = contact.FullName;
                if (fullName != null && fullName.ContainsWholeWord(name, ignoreCase: true))
                {
                    return true;
                }
            }

            Outlook.ExchangeUser exchageUser = addressEntry.GetExchangeUser();
            if (exchageUser != null)
            {
                string alias = exchageUser.Alias;
                if (alias.ContainsWholeWord(name, ignoreCase: true))
                {
                    return true;
                }

                string primarySmptAddress = exchageUser.PrimarySmtpAddress;
                if (primarySmptAddress.ContainsWholeWord(name, ignoreCase: true))
                {
                    return true;
                }
            }

            return false;
        }

        private List<string> ExtractRecipientsFromLine(string line)
        {
            line = line.Trim().Trim('(', ')', '[', ']');

            string[] starters = new[]
            {
                "+",
                "Added",
                "Adding"
            };

            foreach (string starter in starters)
            {
                if (line.StartsWith(starter, StringComparison.OrdinalIgnoreCase))
                {
                    // Skip the starter
                    return ExtractNames(line.Substring(starter.Length));
                }
            }

            return new List<string>();
        }

        private List<string> ExtractNames(string line)
        {
            Regex nameRegex = new Regex("^[a-zA-Z ]*$");
            List<string> entries = line.Split(',', ';')
                .Select(e => e.Trim().TrimEnd('.').Trim())
                .Where(e => nameRegex.IsMatch(e))
                .ToList();
            return entries;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
