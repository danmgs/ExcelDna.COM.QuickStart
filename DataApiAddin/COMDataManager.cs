using DataApi.Core;
using ExcelDna.ComInterop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DataApiAddin
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class COMDataManager
    {
        public string GetHello(string suffix)
        {
            DataManager DataManager = new DataManager();
            return DataManager.GetHello(suffix);
        }
    }
}
