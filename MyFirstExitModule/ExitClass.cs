using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using CERTEXITLib;
using CERTCLILib;
using Microsoft.Win32;

// Kudos to https://nach0focht.wordpress.com/2014/01/05/exit-modules/

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

        private string OutputDirectory = @"C:";

        [DllImport(@"oleaut32.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall)]
        static extern Int32 VariantClear(IntPtr pvarg);

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
            RegistryKey configRegistryKey = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services\CertSvc\Configuration\" + strConfig + @"\ExitModules\MyFirstExitModule");
            if (configRegistryKey != null)
            {
                OutputDirectory = (string)configRegistryKey.GetValue("OutputDirectory");
            }

            return (Int32)(
                ExitEvents.CertIssued |
                ExitEvents.CertPending |
                ExitEvents.CertDenied |
                ExitEvents.CertRevoked |
                ExitEvents.CertRetrievePending | 
                ExitEvents.CRLIssued |
                ExitEvents.Shutdown
                );
        }

        public void Notify(int ExitEvent, int Context)
        {
            string strNotify = null;

            switch (ExitEvent)
            {
                case (int)ExitEvents.CertDenied:
                    break;
                case (int)ExitEvents.CertIssued:

                    // https://blogs.msdn.microsoft.com/alejacma/2008/08/04/how-to-modify-an-interop-assembly-to-change-the-return-type-of-a-method-vb-net/
                    Object variantObject;
                    IntPtr variantObjectPtr;
                    variantObjectPtr = Marshal.AllocHGlobal(2048);

                    CCertServerExit CertServer = new CCertServerExit();

                    // CA properties must be queried before the SetContext Method

                    // Must be called before querying Certificate properties
                    // https://docs.microsoft.com/en-us/windows/win32/api/certif/nf-certif-icertserverexit-setcontext
                    CertServer.SetContext(Context);

                    // Retrieving Certificate Properties
                    // https://docs.microsoft.com/en-us/windows/win32/api/certif/nf-certif-icertserverexit-getcertificateproperty

                    // Get VARIANT containing certificate bytes
                    CertServer.GetCertificateProperty("RawCertificate", PROPTYPE_BINARY, variantObjectPtr);

                    // Read ANSI BSTR information from the VARIANT as we know RawCertificate property
                    // is ANSI BSTR. Please note that the below code is written based on how the
                    // VARIANT structure looks like in C/C++

                    IntPtr bstrPtr;
                    bstrPtr = Marshal.ReadIntPtr(variantObjectPtr, 8);
                    int bstrLen;
                    bstrLen = Marshal.ReadInt32(bstrPtr, -4);
                    byte[] RawCertificate = new byte[bstrLen];
                    Marshal.Copy(bstrPtr, RawCertificate, 0, bstrLen);
                    VariantClear(variantObjectPtr);

                    // Get VARIANT containing RequestId
                    CertServer.GetCertificateProperty("RequestId", PROPTYPE_LONG, variantObjectPtr);
                    variantObject = Marshal.GetObjectForNativeVariant(variantObjectPtr);
                    int RequestId = (int)variantObject;
                    VariantClear(variantObjectPtr);

                    /*

                    // Get VARIANT containing Serial Number
                    CertServer.GetCertificateProperty("SerialNumber", PROPTYPE_STRING, variantObjectPtr);
                    variantObject = Marshal.GetObjectForNativeVariant(variantObjectPtr);
                    string SerialNumber = (String)variantObject;
                    VariantClear(variantObjectPtr);

                    // Get VARIANT containing Distinguished Name
                    CertServer.GetCertificateProperty("DistinguishedName", PROPTYPE_STRING, variantObjectPtr);
                    variantObject = Marshal.GetObjectForNativeVariant(variantObjectPtr);
                    string DistinguishedName = (String)variantObject;
                    VariantClear(variantObjectPtr);

                    // Get VARIANT containing Not Before date
                    CertServer.GetCertificateProperty("NotBefore", PROPTYPE_DATE, variantObjectPtr);
                    variantObject = Marshal.GetObjectForNativeVariant(variantObjectPtr);
                    DateTime NotBefore = (DateTime)variantObject;
                    VariantClear(variantObjectPtr);

                    // Get VARIANT containing Not After date
                    CertServer.GetCertificateProperty("NotAfter", PROPTYPE_DATE, variantObjectPtr);
                    variantObject = Marshal.GetObjectForNativeVariant(variantObjectPtr);
                    DateTime NotAfter = (DateTime)variantObject;
                    VariantClear(variantObjectPtr);

                    */


                    // Free the memory block allocated to hold VARIANT
                    Marshal.FreeHGlobal(variantObjectPtr);

                    // Write the Certificate to File
                    if (OutputDirectory != null)
                    {
                        //System.IO.File.WriteAllBytes(@"C:\" + RequestId.ToString() + ".cer", RawCertificate);
                        System.IO.File.WriteAllText(OutputDirectory + @"\" + RequestId.ToString() + ".cer", Convert.ToBase64String(RawCertificate));
                    }

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
