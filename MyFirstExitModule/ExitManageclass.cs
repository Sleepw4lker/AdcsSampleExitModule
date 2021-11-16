using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using CERTEXITLib;

namespace MyFirstExitModule
{
    [ComVisible(true)]  // Expose to COM
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("MyFirstExitModule.ExitManage")] //this is the progId that will be used when registering the dll that you can look up in the registry
    [Guid("5f036df3-c2c4-42a0-918c-24a376828e57")]  //this is the GUID the dll's will be registered under, it should be unique for each exit module
    public class ExitManage : CERTEXITLib.ICertManageModule, CERTEXITLib.CCertManageExitModule
    {

        #region Constructors

        public ExitManage()
        {
        }

        #endregion

        #region ICertManageModule Members

        public void Configure(string strConfig, string strStorageLocation, int Flags)
        {
        }

        public object GetProperty(string strConfig, string strStorageLocation, string strPropertyName, int Flags)
        {
            string strMsg = null;

            switch (strPropertyName) // Each of these is required
            {
                // The name visible when adding the exit module
                case "Name":
                    strMsg = "MyFirstExitModule"; 
                    break;

                // Description visible from AD CS exit module tab
                case "Description":
                    strMsg = "My first Exit Module";
                    break;

                case "Copyright":
                    strMsg = "Copyright © 2021 Uwe Gradenegger";
                    break;

                case "File Version":
                    strMsg = "0.0.1"; // Insert Values from the AssemblyInfo here
                    break;

                case "Product Version":
                    strMsg = "0.0.1"; // Insert Values from the AssemblyInfo here
                    break;

                default:
                    strMsg = "Unknown Property: " + strPropertyName;
                    break;
            }

            return (object)strMsg;
        }

        public void SetProperty(string strConfig, string strStorageLocation, string strPropertyName, int Flags, ref object pvarProperty)
        {
            // This method is intended for future functionality.
        }

        #endregion

    }
}
