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

    // TODO: i think thish should be distinct from ExitManage class
    [Guid("dc5b174f-076e-4d0e-9066-cdb9f2e6e58c")]  // This is the GUID the dll's will be registered under, it should be unique for each exit module
    public class Exit : CERTEXITLib.ICertExit2, CERTEXITLib.ICertExit
    {
        private string OutputDirectory = null;

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

                    var CertServer = new CCertServerExit();
                    var CertificateInfo = new IssuedCertificate();
                    
                    // Retrieving CA Properties

                    // The context must be zero to read any of these properties. The context is set to zero when the ICertServerExit object is initially created.
                    // It can also be set to zero by invoking the SetContext method.
                    CertServer.SetContext(0);

                    CertificateInfo.Issuer = GetStringCertificateProperty(ref CertServer, "SanitizedCAName");

                    // Must be called before querying Certificate properties
                    // https://docs.microsoft.com/en-us/windows/win32/api/certif/nf-certif-icertserverexit-setcontext
                    CertServer.SetContext(Context);

                    // Retrieving Certificate Properties
                    // https://docs.microsoft.com/en-us/windows/win32/api/certif/nf-certif-icertserverexit-getcertificateproperty

                    CertificateInfo.RawCertificate = Convert.ToBase64String(
                        GetBinaryCertificateProperty(ref CertServer, "RawCertificate"),
                        Base64FormattingOptions.None
                        );

                    CertificateInfo.RequestId = GetLongCertificateProperty(ref CertServer, "RequestId");
                    CertificateInfo.SerialNumber = GetStringCertificateProperty(ref CertServer, "SerialNumber");
                    CertificateInfo.Subject = GetStringCertificateProperty(ref CertServer, "DistinguishedName");
                    CertificateInfo.NotBefore = GetDateCertificateProperty(ref CertServer, "NotBefore");
                    CertificateInfo.NotAfter = GetDateCertificateProperty(ref CertServer, "NotAfter");


                    // Retrieving Request Properties
                    // https://docs.microsoft.com/de-de/windows/win32/api/certif/nf-certif-icertserverexit-getrequestproperty

                    CertificateInfo.RawRequest = Convert.ToBase64String(
                        GetBinaryRequestProperty(ref CertServer, "RawRequest"), 
                        Base64FormattingOptions.None
                        );

                    CertificateInfo.RequesterName = GetStringRequestProperty(ref CertServer, "RequesterName");
                    CertificateInfo.RequestType = GetLongRequestProperty(ref CertServer, "RequestType");

                    // TODO: Can we access the "Name Properties" here, and should we, in Terms of accessing the CA Database...?
                    // https://docs.microsoft.com/en-us/windows/win32/seccrypto/name-properties

                    // Serialize to XML
                    var OutputXmlSerializer = new XmlSerializer(typeof(IssuedCertificate));

                    var OutputFileName = CertificateInfo.RequestId.ToString() + ".xml";

                    /*
                    OutputFileName.Replace("%1", CertificateInfo.RequestId.ToString());
                    OutputFileName.Replace("%2", CertificateInfo.Issuer);
                    OutputFileName.Replace("%3", CertificateInfo.SerialNumber);
                    */

                    using (var OutputStringWriter = new StringWriter())
                    {
                        using (XmlWriter OutputXmlWriter = XmlWriter.Create(OutputStringWriter))
                        {
                            OutputXmlSerializer.Serialize(OutputXmlWriter, CertificateInfo);

                            if (OutputDirectory != null)
                            {
                                System.IO.File.WriteAllText(
                                    OutputDirectory + @"\" + OutputFileName,
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

        #region Property retrieval functions

        [DllImport(@"oleaut32.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall)]
        static extern Int32 VariantClear(IntPtr pvarg);

        private DateTime GetDateCertificateProperty(ref CCertServerExit server, string name)
        {
            IntPtr variantObjectPtr = Marshal.AllocHGlobal(2048);

            try
            {
                server.GetCertificateProperty(name, (int)PropertyType.PROPTYPE_DATE, variantObjectPtr);
                var result = (DateTime)(Marshal.GetObjectForNativeVariant(variantObjectPtr));
                return result;
            }
            catch
            {
                return new DateTime();
            }
            finally
            {
                VariantClear(variantObjectPtr);
                Marshal.FreeHGlobal(variantObjectPtr);
            }
        }
        private string GetStringCertificateProperty(ref CCertServerExit server, string name)
        {
            IntPtr variantObjectPtr = Marshal.AllocHGlobal(2048);

            try
            {
                server.GetCertificateProperty(name, (int)PropertyType.PROPTYPE_STRING, variantObjectPtr);
                var result = (string)(Marshal.GetObjectForNativeVariant(variantObjectPtr));
                return result;
            }
            catch
            {
                return null;
            }
            finally
            {
                VariantClear(variantObjectPtr);
                Marshal.FreeHGlobal(variantObjectPtr);
            }
        }
        private int GetLongCertificateProperty(ref CCertServerExit server, string name)
        {
            IntPtr variantObjectPtr = Marshal.AllocHGlobal(2048);

            try
            {
                server.GetCertificateProperty(name, (int)PropertyType.PROPTYPE_LONG, variantObjectPtr);
                var result = (int)(Marshal.GetObjectForNativeVariant(variantObjectPtr));
                return result;
            }
            catch
            {
                return 0;
            }
            finally
            {
                VariantClear(variantObjectPtr);
                Marshal.FreeHGlobal(variantObjectPtr);
            }
        }
        private Byte[] GetBinaryCertificateProperty(ref CCertServerExit server, string name)
        {
            // https://blogs.msdn.microsoft.com/alejacma/2008/08/04/how-to-modify-an-interop-assembly-to-change-the-return-type-of-a-method-vb-net/
            IntPtr variantObjectPtr = Marshal.AllocHGlobal(2048);

            try
            {
                IntPtr bstrPtr;
                int bstrLen;

                // Get VARIANT containing certificate bytes
                // Read ANSI BSTR information from the VARIANT as we know RawCertificate property is ANSI BSTR.
                server.GetCertificateProperty(name, (int)PropertyType.PROPTYPE_BINARY, variantObjectPtr);
                bstrPtr = Marshal.ReadIntPtr(variantObjectPtr, 8);
                bstrLen = Marshal.ReadInt32(bstrPtr, -4);
                byte[] result = new byte[bstrLen];
                Marshal.Copy(bstrPtr, result, 0, bstrLen);

                return result;
            }
            catch
            {
                return null;
            }
            finally
            {
                VariantClear(variantObjectPtr);
                Marshal.FreeHGlobal(variantObjectPtr);
            }
        }

        private DateTime GetDateRequestProperty(ref CCertServerExit server, string name)
        {
            IntPtr variantObjectPtr = Marshal.AllocHGlobal(2048);

            try
            {
                server.GetRequestProperty(name, (int)PropertyType.PROPTYPE_DATE, variantObjectPtr);
                var result = (DateTime)(Marshal.GetObjectForNativeVariant(variantObjectPtr));
                return result;
            }
            catch
            {
                return new DateTime();
            }
            finally
            {
                VariantClear(variantObjectPtr);
                Marshal.FreeHGlobal(variantObjectPtr);
            }
        }
        private string GetStringRequestProperty(ref CCertServerExit server, string name)
        {
            IntPtr variantObjectPtr = Marshal.AllocHGlobal(2048);

            try
            {
                server.GetRequestProperty(name, (int)PropertyType.PROPTYPE_STRING, variantObjectPtr);
                var result = (string)(Marshal.GetObjectForNativeVariant(variantObjectPtr));
                return result;
            }
            catch
            {
                return null;
            }
            finally
            {
                VariantClear(variantObjectPtr);
                Marshal.FreeHGlobal(variantObjectPtr);
            }

        }
        private int GetLongRequestProperty(ref CCertServerExit server, string name)
        {
            IntPtr variantObjectPtr = Marshal.AllocHGlobal(2048);

            try
            {
                server.GetRequestProperty(name, (int)PropertyType.PROPTYPE_LONG, variantObjectPtr);
                var result = (int)(Marshal.GetObjectForNativeVariant(variantObjectPtr));
                return result;
            }
            catch
            {
                return 0;
            }
            finally
            {
                VariantClear(variantObjectPtr);
                Marshal.FreeHGlobal(variantObjectPtr);
            }
        }
        private Byte[] GetBinaryRequestProperty(ref CCertServerExit server, string name)
        {
            // https://blogs.msdn.microsoft.com/alejacma/2008/08/04/how-to-modify-an-interop-assembly-to-change-the-return-type-of-a-method-vb-net/
            IntPtr variantObjectPtr = Marshal.AllocHGlobal(2048);

            try
            {
                IntPtr bstrPtr;
                int bstrLen;

                // Get VARIANT containing certificate bytes
                // Read ANSI BSTR information from the VARIANT as we know RawCertificate property is ANSI BSTR.
                server.GetRequestProperty(name, (int)PropertyType.PROPTYPE_BINARY, variantObjectPtr);
                bstrPtr = Marshal.ReadIntPtr(variantObjectPtr, 8);
                bstrLen = Marshal.ReadInt32(bstrPtr, -4);
                byte[] result = new byte[bstrLen];
                Marshal.Copy(bstrPtr, result, 0, bstrLen);
                return result;
            }
            catch
            {
                return null;
            }
            finally
            {
                VariantClear(variantObjectPtr);
                Marshal.FreeHGlobal(variantObjectPtr);
            }
        }

        #endregion 

    }
}
