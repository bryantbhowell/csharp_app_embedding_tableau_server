using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.ComponentModel;
using System.Runtime.InteropServices;



namespace Embedded_Tableau_Example
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }

    /*
     * ExtendedWebBrowser class allows for additional events to be bubbled up out of the ActiveX control
     * We need NewWindow3 so we can reroute any new windows that Tableau Server page tries to open up by default
     * 
     */
    public class ExtendedWebBrowser : System.Windows.Forms.WebBrowser
    {
        AxHost.ConnectionPointCookie cookie;
        WebBrowserExtendedEvents events;

        protected override void CreateSink()
        {
            base.CreateSink();
            events = new WebBrowserExtendedEvents(this);
            cookie = new AxHost.ConnectionPointCookie(this.ActiveXInstance, events, typeof(DWebBrowserEvents2));
        }

        public object Application
        {
            get
            {
                IWebBrowser2 axWebBrowser = this.ActiveXInstance as IWebBrowser2;
                if (axWebBrowser != null)
                {
                    return axWebBrowser.Application;
                }
                else
                    return null;
            }
        }

        protected override void DetachSink()
        {
            if (null != cookie)
            {
                cookie.Disconnect();
                cookie = null;
            }
            base.DetachSink();
        }

        //This new event will fire for the NewWindow2
        public event EventHandler<NewWindow2EventArgs> NewWindow2;

        protected void OnNewWindow2(ref object ppDisp, ref bool cancel)
        {
            EventHandler<NewWindow2EventArgs> h = NewWindow2;
            NewWindow2EventArgs args = new NewWindow2EventArgs(ref ppDisp, ref cancel);
            if (null != h)
            {
                h(this, args);
            }
            //Pass the cancellation chosen back out to the events
            //Pass the ppDisp chosen back out to the events
            cancel = args.Cancel;
            ppDisp = args.PPDisp;
        }

        public event EventHandler<NewWindow3EventArgs> NewWindow3;

        protected void OnNewWindow3(ref object ppDisp, ref bool cancel, long dwFlags, string bstrUrlContext, string bstrUrl)
        {
            EventHandler<NewWindow3EventArgs> h = NewWindow3;
            NewWindow3EventArgs args = new NewWindow3EventArgs(ref ppDisp, ref cancel, dwFlags, bstrUrlContext, bstrUrl );
            if (null != h)
            {
                h(this, args);
            }
            //Pass the cancellation chosen back out to the events
            //Pass the ppDisp chosen back out to the events
            cancel = args.Cancel;
            ppDisp = args.PPDisp;
            bstrUrlContext = args.BStrURLContext;
            bstrUrl = args.BStrURL;


        }

 
        //This new event will fire for the DocumentComplete
        public event EventHandler<DocumentCompleteEventArgs> DocumentComplete;

        protected void OnDocumentComplete(object ppDisp, object url)
        {
            EventHandler<DocumentCompleteEventArgs> h = DocumentComplete;
            DocumentCompleteEventArgs args = new DocumentCompleteEventArgs(ppDisp, url);
            if (null != h)
            {
                h(this, args);
            }
            //Pass the ppDisp chosen back out to the events
            ppDisp = args.PPDisp;
            //I think url is readonly
        }

        //This new event will fire for the DocumentComplete
        public event EventHandler<CommandStateChangeEventArgs> CommandStateChange;

        protected void OnCommandStateChange(long command, ref bool enable)
        {
            EventHandler<CommandStateChangeEventArgs> h = CommandStateChange;
            CommandStateChangeEventArgs args = new CommandStateChangeEventArgs(command, ref enable);
            if (null != h)
            {
                h(this, args);
            }
        }


        //This class will capture events from the WebBrowser
        public class WebBrowserExtendedEvents : System.Runtime.InteropServices.StandardOleMarshalObject, DWebBrowserEvents2
        {
            ExtendedWebBrowser _Browser;
            public WebBrowserExtendedEvents(ExtendedWebBrowser browser)
            { _Browser = browser; }

            //Implement whichever events you wish
            public void NewWindow2(ref object pDisp, ref bool cancel)
            {
                _Browser.OnNewWindow2(ref pDisp, ref cancel);
            }

            public void NewWindow3(ref object pDisp, ref bool cancel, uint dwFlags, string bstrUrlContext, string bstrUrl)
            {
                _Browser.OnNewWindow3(ref pDisp, ref cancel, dwFlags, bstrUrlContext, bstrUrl);
            }

            //Implement whichever events you wish
            public void DocumentComplete(object pDisp, ref object url)
            {
                _Browser.OnDocumentComplete(pDisp, url);
            }

            //Implement whichever events you wish
            public void CommandStateChange(long command, bool enable)
            {
                _Browser.OnCommandStateChange(command, ref enable);
            }


        }
        [ComImport, Guid("34A715A0-6587-11D0-924A-0020AFC7AC4D"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType(TypeLibTypeFlags.FHidden)]
        public interface DWebBrowserEvents2
        {
            [DispId(0x69)]
            void CommandStateChange([In] long command, [In] bool enable);
            [DispId(0x103)]
            void DocumentComplete([In, MarshalAs(UnmanagedType.IDispatch)] object pDisp, [In] ref object URL);
            [DispId(0xfb)]
            void NewWindow2([In, Out, MarshalAs(UnmanagedType.IDispatch)] ref object pDisp, [In, Out] ref bool cancel);
            [DispId(0x111)]
            void NewWindow3([In, Out, MarshalAs(UnmanagedType.IDispatch)] ref object pDisp, [In, Out, MarshalAs(UnmanagedType.VariantBool)] ref bool Cancel, uint dwFlags, [MarshalAs(UnmanagedType.BStr)] string bstrUrlContext, [MarshalAs(UnmanagedType.BStr)] string bstrUrl);
        }

        [ComImport, Guid("D30C1661-CDAF-11d0-8A3E-00C04FC9E26E"), TypeLibType(TypeLibTypeFlags.FOleAutomation | TypeLibTypeFlags.FDual | TypeLibTypeFlags.FHidden)]
        public interface IWebBrowser2
        {
            [DispId(100)]
            void GoBack();
            [DispId(0x65)]
            void GoForward();
            [DispId(0x66)]
            void GoHome();
            [DispId(0x67)]
            void GoSearch();
            [DispId(0x68)]
            void Navigate([In] string Url, [In] ref object flags, [In] ref object targetFrameName, [In] ref object postData, [In] ref object headers);
            [DispId(-550)]
            void Refresh();
            [DispId(0x69)]
            void Refresh2([In] ref object level);
            [DispId(0x6a)]
            void Stop();
            [DispId(200)]
            object Application { [return: MarshalAs(UnmanagedType.IDispatch)] get; }
            [DispId(0xc9)]
            object Parent { [return: MarshalAs(UnmanagedType.IDispatch)] get; }
            [DispId(0xca)]
            object Container { [return: MarshalAs(UnmanagedType.IDispatch)] get; }
            [DispId(0xcb)]
            object Document { [return: MarshalAs(UnmanagedType.IDispatch)] get; }
            [DispId(0xcc)]
            bool TopLevelContainer { get; }
            [DispId(0xcd)]
            string Type { get; }
            [DispId(0xce)]
            int Left { get; set; }
            [DispId(0xcf)]
            int Top { get; set; }
            [DispId(0xd0)]
            int Width { get; set; }
            [DispId(0xd1)]
            int Height { get; set; }
            [DispId(210)]
            string LocationName { get; }
            [DispId(0xd3)]
            string LocationURL { get; }
            [DispId(0xd4)]
            bool Busy { get; }
            [DispId(300)]
            void Quit();
            [DispId(0x12d)]
            void ClientToWindow(out int pcx, out int pcy);
            [DispId(0x12e)]
            void PutProperty([In] string property, [In] object vtValue);
            [DispId(0x12f)]
            object GetProperty([In] string property);
            [DispId(0)]
            string Name { get; }
            [DispId(-515)]
            int HWND { get; }
            [DispId(400)]
            string FullName { get; }
            [DispId(0x191)]
            string Path { get; }
            [DispId(0x192)]
            bool Visible { get; set; }
            [DispId(0x193)]
            bool StatusBar { get; set; }
            [DispId(0x194)]
            string StatusText { get; set; }
            [DispId(0x195)]
            int ToolBar { get; set; }
            [DispId(0x196)]
            bool MenuBar { get; set; }
            [DispId(0x197)]
            bool FullScreen { get; set; }
            [DispId(500)]
            void Navigate2([In] ref object URL, [In] ref object flags, [In] ref object targetFrameName, [In] ref object postData, [In] ref object headers);
            [DispId(0x1f7)]
            void ShowBrowserBar([In] ref object pvaClsid, [In] ref object pvarShow, [In] ref object pvarSize);
            [DispId(-525)]
            WebBrowserReadyState ReadyState { get; }
            [DispId(550)]
            bool Offline { get; set; }
            [DispId(0x227)]
            bool Silent { get; set; }
            [DispId(0x228)]
            bool RegisterAsBrowser { get; set; }
            [DispId(0x229)]
            bool RegisterAsDropTarget { get; set; }
            [DispId(0x22a)]
            bool TheaterMode { get; set; }
            [DispId(0x22b)]
            bool AddressBar { get; set; }
            [DispId(0x22c)]
            bool Resizable { get; set; }
        }

    }

    public class NewWindow2EventArgs : CancelEventArgs
    {
        object ppDisp;

        public object PPDisp
        {
            get { return ppDisp; }
            set { ppDisp = value; }
        }

        public NewWindow2EventArgs(ref object ppDisp, ref bool cancel)
            : base()
        {
            this.ppDisp = ppDisp;
            this.Cancel = cancel;
        }
    }

    public class NewWindow3EventArgs : CancelEventArgs
    {
        object ppDisp;
        public string BStrURLContext;
        public string BStrURL;

        public object PPDisp
        {
            get { return ppDisp; }
            set { ppDisp = value; }

        }

        public NewWindow3EventArgs(ref object ppDisp, ref bool cancel, long dwFlags, string bstrUrlContext, string bstrUrl)
            : base()
        {
            this.ppDisp = ppDisp;
            this.Cancel = cancel;
            this.BStrURLContext = bstrUrlContext;
            this.BStrURL = bstrUrl;
        }
    }


    public class DocumentCompleteEventArgs : EventArgs
    {
        private object ppDisp;
        private object url;

        public object PPDisp
        {
            get { return ppDisp; }
            set { ppDisp = value; }
        }

        public object Url
        {
            get { return url; }
            set { url = value; }
        }

        public DocumentCompleteEventArgs(object ppDisp, object url)
        {
            this.ppDisp = ppDisp;
            this.url = url;
        }
    }

    public class CommandStateChangeEventArgs : EventArgs
    {
        private long command;
        private bool enable;
        public CommandStateChangeEventArgs(long command, ref bool enable)
        {
            this.command = command;
            this.enable = enable;
        }

        public long Command
        {
            get { return command; }
            set { command = value; }
        }
    }
}
