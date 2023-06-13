using System;

namespace OfficeAutomation.Outlook
{
    /// <summary>
    /// Class meant for sending emails via interaction with Microsoft Outlook desktop application. It can either interact with a running 
    /// Microsoft Outlook application, or it can start a new instance of Microsoft Outlook application. 
    /// </summary>
    public class EmailSender
    {
        protected Microsoft.Office.Interop.Outlook.Application _outlookApplication = null;

        protected virtual Microsoft.Office.Interop.Outlook.Application GetOutlookApplication()
        {
            if (this._outlookApplication == null)
            {
                this._outlookApplication = ApplicationHandler.GetOutlook();
            }

            return this._outlookApplication;
        }

        /// <summary>
        /// Method sends the provided <paramref name="emailMessage"/> using Microsoft Outlook application. It uses either an existing 
        /// and running instance of Microsoft Outlook application, or it spawns a new one. 
        /// </summary>
        /// 
        /// <param name="emailMessage">Email message with all details to be sent.</param>
        /// 
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="emailMessage"/> is null.</exception>
        public virtual void SendEmail(EmailMessage emailMessage)
        {
            if (emailMessage == null)
            {
                throw new ArgumentNullException(nameof(emailMessage));
            }

            try
            {
                Microsoft.Office.Interop.Outlook.Application application = this.GetOutlookApplication();

                Microsoft.Office.Interop.Outlook.MailItem newMailItem = (Microsoft.Office.Interop.Outlook.MailItem)application.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);

                newMailItem.To = string.Join("; ", emailMessage.Receivers.ToArray());
                newMailItem.CC = string.Join("; ", emailMessage.CCReceivers.ToArray());
                newMailItem.BCC = string.Join("; ", emailMessage.BCCReceivers.ToArray());

                newMailItem.Subject = emailMessage.Subject;

                if (!string.IsNullOrWhiteSpace(emailMessage.PlaintextBody))
                {
                    newMailItem.Body = emailMessage.PlaintextBody;
                }
                else if (!string.IsNullOrWhiteSpace(emailMessage.HtmlBody))
                {
                    newMailItem.HTMLBody = emailMessage.HtmlBody;
                }

                Microsoft.Office.Interop.Outlook.Accounts accounts = application.Session.Accounts;
                foreach (Microsoft.Office.Interop.Outlook.Account account in accounts)
                {
                    // When the e-mail address matches, send the mail. 
                    if (string.Equals(account.SmtpAddress, emailMessage.SendFromEmailAddress, StringComparison.OrdinalIgnoreCase))
                    {
                        newMailItem.SendUsingAccount = account;
                        newMailItem.Save();
                        newMailItem.Send();
                        break;
                    }
                }
            }
            catch (System.Exception)
            {
                // Any exception handling logic can be added here. 
                // Rethrowing the exception here. 

                throw;
            }
        }

        /// <summary>
        /// Constructor creates a new instance of <see cref="EmailSender"/> class.
        /// </summary>
        public EmailSender()
        {
        }
    }
}
