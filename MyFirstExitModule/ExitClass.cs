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
        private string OutputDirectory = null;

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

        internal enum PropertyType : int
        {
            PROPTYPE_LONG = 1,
            PROPTYPE_DATE = 2,
            PROPTYPE_BINARY = 3,
            PROPTYPE_STRING = 4,
            PROPTYPE_ANSI = 5
        }

        internal enum RequestType : int
        {
            CR_IN_PKCS10 = 0x100,
            CR_IN_KEYGEN = 0x200,
            CR_IN_PKCS7 = 0x300
        }

        public class IssuedCertificate
        {
            public int RequestId { get; set; } = 0;
            public string RequesterName { get; set; } = null;
            public string Subject { get; set; } = null;
            public string Issuer { get; set; } = null;
            public string SerialNumber { get; set; } = null;
            public DateTime NotBefore { get; set; }
            public DateTime NotAfter { get; set; }
            public string RawCertificate { get; set; } = null;
            public int RequestType { get; set; }
            public string RawRequest { get; set; } = null;
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
                ExitEvents.CertIssued
            );
        }

        public void Notify(int ExitEvent, int Context)
        {

            switch (ExitEvent)
            {
                case (int)ExitEvents.CertIssued:

                    // https://blogs.msdn.microsoft.com/alejacma/2008/08/04/how-to-modify-an-interop-assembly-to-change-the-return-type-of-a-method-vb-net/
                    IntPtr variantObjectPtr = Marshal.AllocHGlobal(2048);
                    IntPtr bstrPtr;
                    int bstrLen;

                    var CertServer = new CCertServerExit();
                    var CertificateInfo = new IssuedCertificate();
                    
                    // Retrieving CA Properties

                    // The context must be zero to read any of these properties. The context is set to zero when the ICertServerExit object is initially created.
                    // It can also be set to zero by invoking the SetContext method.
                    CertServer.SetContext(0);

                    CertServer.GetCertificateProperty("SanitizedCAName", (int)PropertyType.PROPTYPE_STRING, variantObjectPtr);
                    CertificateInfo.Issuer = (String)(Marshal.GetObjectForNativeVariant(variantObjectPtr));
                    VariantClear(variantObjectPtr);

                    // Must be called before querying Certificate properties
                    // https://docs.microsoft.com/en-us/windows/win32/api/certif/nf-certif-icertserverexit-setcontext
                    CertServer.SetContext(Context);

                    // Retrieving Certificate Properties
                    // https://docs.microsoft.com/en-us/windows/win32/api/certif/nf-certif-icertserverexit-getcertificateproperty

                    // Get VARIANT containing certificate bytes
                    // Read ANSI BSTR information from the VARIANT as we know RawCertificate property is ANSI BSTR.
                    CertServer.GetCertificateProperty("RawCertificate", (int)PropertyType.PROPTYPE_BINARY, variantObjectPtr);
                    bstrPtr = Marshal.ReadIntPtr(variantObjectPtr, 8);                    
                    bstrLen = Marshal.ReadInt32(bstrPtr, -4);
                    byte[] RawCertificate = new byte[bstrLen];
                    Marshal.Copy(bstrPtr, RawCertificate, 0, bstrLen);
                    CertificateInfo.RawCertificate = Convert.ToBase64String(RawCertificate, Base64FormattingOptions.None);
                    VariantClear(variantObjectPtr);

                    CertServer.GetCertificateProperty("RequestId", (int)PropertyType.PROPTYPE_LONG, variantObjectPtr);
                    CertificateInfo.RequestId = (int)(Marshal.GetObjectForNativeVariant(variantObjectPtr));
                    VariantClear(variantObjectPtr);

                    CertServer.GetCertificateProperty("SerialNumber", (int)PropertyType.PROPTYPE_STRING, variantObjectPtr);
                    CertificateInfo.SerialNumber = (String)(Marshal.GetObjectForNativeVariant(variantObjectPtr));
                    VariantClear(variantObjectPtr);

                    CertServer.GetCertificateProperty("DistinguishedName", (int)PropertyType.PROPTYPE_STRING, variantObjectPtr);
                    CertificateInfo.Subject = (String)(Marshal.GetObjectForNativeVariant(variantObjectPtr));
                    VariantClear(variantObjectPtr);

                    CertServer.GetCertificateProperty("NotBefore", (int)PropertyType.PROPTYPE_DATE, variantObjectPtr);
                    CertificateInfo.NotBefore = (DateTime)(Marshal.GetObjectForNativeVariant(variantObjectPtr));
                    VariantClear(variantObjectPtr);

                    CertServer.GetCertificateProperty("NotAfter", (int)PropertyType.PROPTYPE_DATE, variantObjectPtr);
                    CertificateInfo.NotAfter = (DateTime)(Marshal.GetObjectForNativeVariant(variantObjectPtr));
                    VariantClear(variantObjectPtr);

                    // Retrieving Request Properties
                    // https://docs.microsoft.com/de-de/windows/win32/api/certif/nf-certif-icertserverexit-getrequestproperty

                    CertServer.GetRequestProperty("RawRequest", (int)PropertyType.PROPTYPE_BINARY, variantObjectPtr);
                    bstrPtr = Marshal.ReadIntPtr(variantObjectPtr, 8);
                    bstrLen = Marshal.ReadInt32(bstrPtr, -4);
                    byte[] RawRequest = new byte[bstrLen];
                    Marshal.Copy(bstrPtr, RawRequest, 0, bstrLen);
                    CertificateInfo.RawRequest = Convert.ToBase64String(RawRequest, Base64FormattingOptions.None);
                    VariantClear(variantObjectPtr);

                    CertServer.GetRequestProperty("RequesterName", (int)PropertyType.PROPTYPE_STRING, variantObjectPtr);
                    CertificateInfo.RequesterName = (String)(Marshal.GetObjectForNativeVariant(variantObjectPtr));
                    VariantClear(variantObjectPtr);

                    CertServer.GetRequestProperty("RequestType", (int)PropertyType.PROPTYPE_LONG, variantObjectPtr);
                    CertificateInfo.RequestType = (int)(Marshal.GetObjectForNativeVariant(variantObjectPtr));
                    VariantClear(variantObjectPtr);

                    // TODO: Can we access the "Name Properties" here, and should we, in Terms of accessing the CA Database...?
                    // https://docs.microsoft.com/en-us/windows/win32/seccrypto/name-properties

                    // We're done. Free the memory block allocated to hold VARIANT.
                    Marshal.FreeHGlobal(variantObjectPtr);

                    // Serialize to XML
                    var OutputXmlSerializer = new XmlSerializer(typeof(IssuedCertificate));

                    using (var OutputStringWriter = new StringWriter())
                    {
                        using (XmlWriter OutputXmlWriter = XmlWriter.Create(OutputStringWriter))
                        {
                            OutputXmlSerializer.Serialize(OutputXmlWriter, CertificateInfo);

                            if (OutputDirectory != null)
                            {
                                System.IO.File.WriteAllText(
                                    OutputDirectory + @"\" + CertificateInfo.RequestId.ToString() + ".xml",
                                    PrettyXml(OutputStringWriter.ToString())
                                    );
                            }
                        }
                    }

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
