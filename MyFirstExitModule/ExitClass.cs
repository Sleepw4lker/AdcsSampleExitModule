using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using CERTEXITLib;
using CERTCLILib;

namespace MyFirstExitModule
{
    [ComVisible(true)]  //expose to COM
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("MyFirstExitModule.Exit")] //this is the progId that will be used when registering the dll that you can look up in the registry
    [Guid("5f036df3-c2c4-42a0-918c-24a376828e57")]  //this is the GUID the dll's will be registered under, it should be unique for each exit module
    public class Exit : CERTEXITLib.ICertExit2, CERTEXITLib.ICertExit
    {
        //enums needed when calling the GetCertificiateProperty method
        private static int PROPTYPE_LONG = 1;
        private static int PROPTYPE_DATE = 2;
        private static int PROPTYPE_BINARY = 3;
        private static int PROPTYPE_STRING = 4;
        private static int PROPTYPE_ANSI = 5;

        #region Constructors

        public Exit()
        {
        }

        #endregion

        #region ICertExit Members

        internal enum ExitEvents : int
        {
            Invalid = 0x0,
            CertIssued = 0x1,
            CertPending = 0x2,
            CertDenied = 0x4,
            CertRevoked = 0x8,
            CertRetrievePending = 0x10,
            CRLIssued = 0x20,
            Shutdown = 0x40
        }

        //method must return a description string when the server engine requests it
        public string GetDescription()
        {
            return "My First Exit Module";
        }

        public int Initialize(string strConfig)
        {

            return (Int32)(ExitEvents.CertIssued |
                            ExitEvents.CertPending | ExitEvents.CertDenied
                            | ExitEvents.CertRevoked |
                            ExitEvents.CertRetrievePending
                            | ExitEvents.CRLIssued |
                        ExitEvents.Shutdown);
        }

        public void Notify(int ExitEvent, int Context)
        {
            string strNotify = null;

            switch (ExitEvent)
            {
                case (int)ExitEvents.CertDenied:
                    break;
                case (int)ExitEvents.CertIssued:

                    CCertServerExit CertServer = new CCertServerExit();
                    CertServer.SetContext(Context);

                    string l_strSerialNumber = CertServer.GetCertificateProperty("SerialNumber", PROPTYPE_STRING);
                    System.IO.File.WriteAllText(@"C:\Serial.txt", l_strSerialNumber);

                    break;
                case (int)ExitEvents.CertPending:
                    break;
                case (int)ExitEvents.CertRetrievePending:
                    break;
                case (int)ExitEvents.CertRevoked:
                    break;
                case (int)ExitEvents.CRLIssued:
                    break;
                case (int)ExitEvents.Invalid:
                    break;
                case (int)ExitEvents.Shutdown:
                    break;
            }
        }

        #endregion

        #region ICertExit2 Members

        public CCertManageExitModule GetManageModule()
        {
            return new ExitManage();
        }

        #endregion

    }
}
