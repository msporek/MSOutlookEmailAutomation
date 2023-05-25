using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace OfficeAutomation.Outlook
{
    /// <summary>
    /// Class meant for collaborating with existing instances of Microsoft Outlook, or spawning new instances of Microsoft 
    /// Outlook process. 
    /// </summary>
    public class ApplicationHandler
    {
        public static readonly string OutlookProcessName = "OUTLOOK";

        public static readonly string NamespaceMAPI = "MAPI";

        public static readonly string OutlookApplicationProgID = "Outlook.Application";

        /// <summary>
        /// Method attempts to get an existing (i.e. running) Microsoft Outlook application. 
        /// </summary>
        /// 
        /// <returns>An instance of <see cref="Microsoft.Office.Interop.Outlook.Application"/> class representing an existing 
        /// (i.e. running) Microsoft Outlook, or null if Outlook is not running.</returns>
        public static Microsoft.Office.Interop.Outlook.Application GetRunningOutlook()
        {
            // Check whether there is an Outlook process running. 
            if (Process.GetProcessesByName(ApplicationHandler.OutlookProcessName).Count() > 0)
            {
                // If so, use the GetActiveObject method to obtain the process and cast it to an Application object. 
                return Marshal.GetActiveObject(ApplicationHandler.OutlookApplicationProgID) as Microsoft.Office.Interop.Outlook.Application;
            }

            return null;
        }

        /// <summary>
        /// Method attempts to get an existing (i.e. running) Microsoft Outlook application, or otherwise to instantiate a 
        /// new Microsoft Outlook process. 
        /// </summary>
        /// 
        /// <returns>An instance of <see cref="Microsoft.Office.Interop.Outlook.Application"/> class representing an existing 
        /// (i.e. running) Microsoft Outlook, or a newly-instantiated Microsoft Outlook application process.</returns>
        public static Microsoft.Office.Interop.Outlook.Application GetOutlook()
        {
            Microsoft.Office.Interop.Outlook.Application application = ApplicationHandler.GetRunningOutlook();
            if (application == null)
            {
                // If not, create a new instance of Outlook and sign in to the default profile. 
                application = new Microsoft.Office.Interop.Outlook.Application();
                Microsoft.Office.Interop.Outlook.NameSpace nameSpace = application.GetNamespace(ApplicationHandler.NamespaceMAPI);
                nameSpace.Logon(string.Empty, string.Empty, Missing.Value, Missing.Value);
                nameSpace = null;
            }

            // Return the Outlook Application object. 
            return application;
        }
    }
}
