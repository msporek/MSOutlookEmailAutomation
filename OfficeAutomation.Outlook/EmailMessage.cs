using System.Collections.Generic;

namespace OfficeAutomation.Outlook
{
    /// <summary>
    /// Class represents details of an email message to be send with the library. The class properties represent all properties that 
    /// the library supports when sending emails. 
    /// </summary>
    public class EmailMessage
    {
        public static EmailMessage FromPlaintextBody(string subject, string plaintextBody, string sendFromEmailAddress, List<string> receivers)
        {
            return new EmailMessage(subject, plaintextBody, null, sendFromEmailAddress, receivers);
        }

        public static EmailMessage FromHtmlBody(string subject, string htmlBody, string sendFromEmailAddress, List<string> receivers)
        {
            return new EmailMessage(subject, null, htmlBody, sendFromEmailAddress, receivers);
        }

        public string Subject { get; private set; }

        public string PlaintextBody { get; private set; }

        public string HtmlBody { get; private set; }

        public string SendFromEmailAddress { get; private set; }

        public List<string> Receivers { get; private set; }

        public List<string> CCReceivers { get; private set; }

        public List<string> BCCReceivers { get; private set; }

        public EmailMessage()
        {
        }

        public EmailMessage(
            string subject, 
            string plaintextBody, 
            string htmlBody, 
            string sendFromEmailAddress, 
            List<string> receivers, 
            List<string> ccReceivers = null, 
            List<string> bccReceivers = null)
            : this()
        {
            this.Subject = subject;

            this.PlaintextBody = plaintextBody;
            this.HtmlBody = htmlBody;

            this.SendFromEmailAddress = sendFromEmailAddress;

            this.Receivers = receivers;
            
            this.CCReceivers = ccReceivers ?? new List<string>();
            this.BCCReceivers = bccReceivers ?? new List<string>();
        }
    }
}
