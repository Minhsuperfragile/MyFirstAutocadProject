using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace MyFirstAutocadProject 
{
    public class Init : Autodesk.AutoCAD.Runtime.IExtensionApplication 
        {
        #region Initialization
        public void Initialize() {
            AcAp.ShowAlertDialog("Loaded this Plugin!");
        }

        public void Terminate() {

        }
        #endregion
    }
}
