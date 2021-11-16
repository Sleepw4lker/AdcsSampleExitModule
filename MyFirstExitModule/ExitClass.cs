using System;
using System.Runtime.InteropServices;
using System.IO;
using System.Xml;
using System.Text;
using System.Xml.Linq;
using System.Xml.Serialization;
using CERTEXITLib;
using CERTCLILib;
using Microsoft.Win32;

// Kudos to https://nach0focht.wordpress.com/2014/01/05/exit-modules/

namespace MyFirstExitModule
{
    [ComVisible(true)]  // Expose to COM
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("MyFirstExitModule.Exit")] // This is the progId that will be used when registering the dll that you can look up in the registry
    [Guid("5f036df3-c2c4-42a0-918c-24a376828e57")]  // This is the GUID the dll's will be registered under, it should be unique for each exit module
    public class Exit : CERTEXITLib.ICertExit2, CERTEXITLib.ICertExit
    {
        // Enums needed when calling the GetCertificiateProperty method
        private static int PROPTYPE_LONG = 1;
        private static int PROPTYPE_DATE = 2;
        private static int PROPTYPE_BINARY = 3;
        private static int PROPTYPE_STRING = 4;
        private static int PROPTYPE_ANSI = 5;

        /*
        private static int CR_IN_PKCS10 = 0x100;
        private static int CR_IN_KEYGEN = 0x200;
        private static int CR_IN_PKCS7 = 0x300;
         * */

        // Build a Logic here
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

        public class IssuedCertificate
        {
            public UInt32 RequestId { get; set; } = 0;
            public string RequesterName { get; set; } = null;
            public int RequestType { get; set; }
            public string RawRequest { get; set; } = null;
            public string SerialNumber { get; set; } = null;
            public DateTime NotBefore { get; set; }
            public DateTime NotAfter { get; set; }
            public string DistinguishedName { get; set; } = null;
            public string RawCertificate { get; set; } = null;
        }

        // https://stackoverflow.com/questions/1123718/format-xml-string-to-print-friendly-xml-string
        private static string PrettyXml(string xml)
        {
            var stringBuilder = new StringBuilder();

            var element = XElement.Parse(xml);

            var settings = new XmlWriterSettings();
            settings.OmitXmlDeclaration = true;
            settings.Indent = true;
            settings.NewLineOnAttributes = true;

            using (var xmlWriter = XmlWriter.Create(stringBuilder, settings))
            {
                element.Save(xmlWriter);
            }

            return stringBuilder.ToString();
        }

        // Method must return a description string when the server engine requests it
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

            // Subscribe to the Events we want to process
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

            switch (ExitEvent)
            {
                case (int)ExitEvents.CertDenied:
                    break;
                case (int)ExitEvents.CertIssued:

                    // https://blogs.msdn.microsoft.com/alejacma/2008/08/04/how-to-modify-an-interop-assembly-to-change-the-return-type-of-a-method-vb-net/
                    Object variantObject;
                    IntPtr variantObjectPtr;
                    IntPtr bstrPtr;
                    int bstrLen;
                    variantObjectPtr = Marshal.AllocHGlobal(2048);

                    CCertServerExit CertServer = new CCertServerExit();

                    // CA properties must be queried before the SetContext Method, or SetContext must be set back to zero

                    // Must be called before querying Certificate properties
                    // https://docs.microsoft.com/en-us/windows/win32/api/certif/nf-certif-icertserverexit-setcontext
                    CertServer.SetContext(Context);

                    // Retrieving Certificate Properties
                    // https://docs.microsoft.com/en-us/windows/win32/api/certif/nf-certif-icertserverexit-getcertificateproperty

                    // Get VARIANT containing certificate bytes
                    // Read ANSI BSTR information from the VARIANT as we know RawCertificate property is ANSI BSTR.
                    // Please note that the below code is written based on how the VARIANT structure looks like in C/C++
                    CertServer.GetCertificateProperty("RawCertificate", PROPTYPE_BINARY, variantObjectPtr);
                    bstrPtr = Marshal.ReadIntPtr(variantObjectPtr, 8);                    
                    bstrLen = Marshal.ReadInt32(bstrPtr, -4);
                    byte[] RawCertificate = new byte[bstrLen];
                    Marshal.Copy(bstrPtr, RawCertificate, 0, bstrLen);
                    VariantClear(variantObjectPtr);

                    CertServer.GetCertificateProperty("RequestId", PROPTYPE_LONG, variantObjectPtr);
                    variantObject = Marshal.GetObjectForNativeVariant(variantObjectPtr);
                    UInt32 RequestId = (UInt32)variantObject;
                    VariantClear(variantObjectPtr);

                    CertServer.GetCertificateProperty("SerialNumber", PROPTYPE_STRING, variantObjectPtr);
                    variantObject = Marshal.GetObjectForNativeVariant(variantObjectPtr);
                    string SerialNumber = (String)variantObject;
                    VariantClear(variantObjectPtr);

                    CertServer.GetCertificateProperty("DistinguishedName", PROPTYPE_STRING, variantObjectPtr);
                    variantObject = Marshal.GetObjectForNativeVariant(variantObjectPtr);
                    string DistinguishedName = (String)variantObject;
                    VariantClear(variantObjectPtr);

                    CertServer.GetCertificateProperty("NotBefore", PROPTYPE_DATE, variantObjectPtr);
                    variantObject = Marshal.GetObjectForNativeVariant(variantObjectPtr);
                    DateTime NotBefore = (DateTime)variantObject;
                    VariantClear(variantObjectPtr);

                    CertServer.GetCertificateProperty("NotAfter", PROPTYPE_DATE, variantObjectPtr);
                    variantObject = Marshal.GetObjectForNativeVariant(variantObjectPtr);
                    DateTime NotAfter = (DateTime)variantObject;
                    VariantClear(variantObjectPtr);

                    // Retrieving Request Properties
                    // https://docs.microsoft.com/de-de/windows/win32/api/certif/nf-certif-icertserverexit-getrequestproperty

                    CertServer.GetRequestProperty("RawRequest", PROPTYPE_BINARY, variantObjectPtr);
                    bstrPtr = Marshal.ReadIntPtr(variantObjectPtr, 8);
                    bstrLen = Marshal.ReadInt32(bstrPtr, -4);
                    byte[] RawRequest = new byte[bstrLen];
                    Marshal.Copy(bstrPtr, RawRequest, 0, bstrLen);
                    VariantClear(variantObjectPtr);
                    
                    CertServer.GetRequestProperty("RequesterName", PROPTYPE_STRING, variantObjectPtr);
                    variantObject = Marshal.GetObjectForNativeVariant(variantObjectPtr);
                    String RequesterName = (String)variantObject;
                    VariantClear(variantObjectPtr);

                    CertServer.GetRequestProperty("RequestType", PROPTYPE_LONG, variantObjectPtr);
                    variantObject = Marshal.GetObjectForNativeVariant(variantObjectPtr);
                    int RequestType = (int)variantObject;
                    VariantClear(variantObjectPtr);

                    // TODO: Can we access the "Name Properties" here, and should we, in Terms of accessing the CA Database...?
                    // https://docs.microsoft.com/en-us/windows/win32/seccrypto/name-properties

                    // We're done. Free the memory block allocated to hold VARIANT.
                    Marshal.FreeHGlobal(variantObjectPtr);

                    // Build the Data Structure for export
                    var CertInfo = new IssuedCertificate
                    {
                        RequestId = RequestId,
                        RequesterName = RequesterName,
                        RequestType = RequestType,
                        RawRequest = Convert.ToBase64String(RawRequest, Base64FormattingOptions.None),
                        SerialNumber = SerialNumber,
                        NotBefore = NotBefore,
                        NotAfter = NotAfter,
                        DistinguishedName = DistinguishedName,
                        RawCertificate = Convert.ToBase64String(RawCertificate, Base64FormattingOptions.None)
                    };

                    // Serialize to XML
                    System.Xml.Serialization.XmlSerializer OutputXmlSerializer = new XmlSerializer(typeof(IssuedCertificate));

                    using (var OutputStringWriter = new StringWriter())
                    {
                        using (XmlWriter OutputXmlWriter = XmlWriter.Create(OutputStringWriter))
                        {
                            OutputXmlSerializer.Serialize(OutputXmlWriter, CertInfo);

                            if (OutputDirectory != null)
                            {
                                System.IO.File.WriteAllText(
                                    OutputDirectory + @"\" + RequestId.ToString() + ".xml",
                                    PrettyXml(OutputStringWriter.ToString())
                                    );
                            }
                        }
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
