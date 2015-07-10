using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Security.Permissions;
using System.Net;
using Microsoft.Win32;
using System.Threading.Tasks;

namespace Embedded_Tableau_Example
{
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    [System.Runtime.InteropServices.ComVisibleAttribute(true)]
    public partial class Form1 : Form
    {
        int tab_count = 1;
        String username = "bhowell"; // Here you can get your username however you'd like
        bool edit_mode = false; // Flag for when we are in web edit mode
        public Form1()
        {
            InitializeComponent();
            SetWebBrowserFeatures();
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            string postData = String.Format("username={0}", username);
            byte[] Post = Encoding.UTF8.GetBytes(postData);
            string AdditionalHeaders = "Content-Type: application/x-www-form-urlencoded";

            var page_location = "http://bhowell-lap:88/portal/c_sharp_embed.php";
            //webBrowser1.Navigate(page_location); // If using AD with automatic logon
            webBrowser1.Navigate(page_location, "_self", Post, AdditionalHeaders);
            webBrowser1.ObjectForScripting = this;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("pdfExport");
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            
            object[] args = { tab_count };

            webBrowser1.Document.InvokeScript("switchTab", args);
            tab_count++; 

        }

        public void Test(String message)
        {
            MessageBox.Show(message, "client code");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("returnJSONString");
        }

        public void retrieveJSON(String json)
        {
            MessageBox.Show(json, "Stringy JSON");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            webBrowser1.Document.InvokeScript("launch_edit");
            edit_mode = true;
        }

        private void webBrowser1_NewWindow(object sender, CancelEventArgs e)
        {

           // e.Cancel = true;
        }

        /*
         * The NewWindow3 object is triggered whenever a page tries to open a new window, which by default opens in a new IE window. 
         * NewWindow3 has property BStrURL which is the location to be loaded in the new window. If you cancel the action, you can
         * capture the location and reroute or choose a different set of actions.
         */
        private void webBrowser1_NewWindow3(object sender, NewWindow3EventArgs e)
        {
            
            //MessageBox.Show("We going somewhere", "Where we going?");
            //MessageBox.Show(e.BStrURL, "Going here");
            if (e.BStrURL.Contains("authoring")) {
                e.Cancel = true;
                MessageBox.Show(e.BStrURL, "Web Authoring Triggered");
                edit_mode = true;
                webBrowser1.Navigate(e.BStrURL);
                
            }
            // This one needs the window to atually show up, maybe place in its own separate tab? But will that maintain session??
            else if (e.BStrURL.Contains("viewData"))
            {
                e.Cancel = true;
                webBrowser1.Navigate(e.BStrURL);
            }
            // For all of these, redirecting the URL back into current window causes the download dialog to appear successfully
            else if (
                        e.BStrURL.Contains("tempfile") ||
                        e.BStrURL.Contains("crosstab") ||
                        e.BStrURL.Contains("csv=true")
                    )
            {
                e.Cancel = true;
                webBrowser1.Navigate(e.BStrURL);
            }
            
        }

        private void webBrowser1_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            //   MessageBox.Show("We going somewhere WITHIN","Where we going?");
            if (edit_mode == true)
            {

                string url_string = e.Url.OriginalString;
                // This is what happens when Done or Save As happens
                if (url_string.Contains("views") ) {
                    MessageBox.Show(url_string, "Going back");
                    edit_mode = false;
                    string postData = String.Format("username={0}&view_location={1}", username, e.Url.OriginalString);
                    byte[] Post = Encoding.UTF8.GetBytes(postData);
                    string AdditionalHeaders = "Content-Type: application/x-www-form-urlencoded";

                    var page_location = "http://bhowell-lap:88/portal/c_sharp_embed.php";
                    //webBrowser1.Navigate(page_location); // If using AD with automatic logon
                    webBrowser1.Navigate(page_location, "_self", Post, AdditionalHeaders);

                }
            }
        }


        /*
         * This section handles forcing the WebBrowser into modern standards compliant mode
         * Without it, you'll get dropped into IE10 compatibility mode and not all the buttons in web edit show up
         * */

        // set WebBrowser features, more info: http://stackoverflow.com/a/18333982/1768303
        static void SetWebBrowserFeatures()
        {
            // don't change the registry if running in-proc inside Visual Studio
            if (LicenseManager.UsageMode != LicenseUsageMode.Runtime)
                return;

            var appName = System.IO.Path.GetFileName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);

            var featureControlRegKey = @"HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\";

            Registry.SetValue(featureControlRegKey + "FEATURE_BROWSER_EMULATION",
                appName, GetBrowserEmulationMode(), RegistryValueKind.DWord);

            // enable the features which are "On" for the full Internet Explorer browser

            Registry.SetValue(featureControlRegKey + "FEATURE_ENABLE_CLIPCHILDREN_OPTIMIZATION",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_AJAX_CONNECTIONEVENTS",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_GPU_RENDERING",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_WEBOC_DOCUMENT_ZOOM",
                appName, 1, RegistryValueKind.DWord);

            Registry.SetValue(featureControlRegKey + "FEATURE_NINPUT_LEGACYMODE",
                appName, 0, RegistryValueKind.DWord);
        }

        static UInt32 GetBrowserEmulationMode()
        {
            int browserVersion = 0;
            using (var ieKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Internet Explorer",
                RegistryKeyPermissionCheck.ReadSubTree,
                System.Security.AccessControl.RegistryRights.QueryValues))
            {
                var version = ieKey.GetValue("svcVersion");
                if (null == version)
                {
                    version = ieKey.GetValue("Version");
                    if (null == version)
                        throw new ApplicationException("Microsoft Internet Explorer is required!");
                }
                int.TryParse(version.ToString().Split('.')[0], out browserVersion);
            }

            if (browserVersion < 8)
            {
                throw new ApplicationException("Unsupported version of Microsoft Internet Explorer!");
            }

            UInt32 mode = 11000; // Internet Explorer 11. Webpages containing standards-based !DOCTYPE directives are displayed in IE11 Standards mode. 

            switch (browserVersion)
            {
                case 7:
                    mode = 7000; // Webpages containing standards-based !DOCTYPE directives are displayed in IE7 Standards mode. 
                    break;
                case 8:
                    mode = 8000; // Webpages containing standards-based !DOCTYPE directives are displayed in IE8 mode. 
                    break;
                case 9:
                    mode = 9000; // Internet Explorer 9. Webpages containing standards-based !DOCTYPE directives are displayed in IE9 mode.                    
                    break;
                case 10:
                    mode = 10000; // Internet Explorer 10.
                    break;
            }

            return mode;
        }


    }
}
