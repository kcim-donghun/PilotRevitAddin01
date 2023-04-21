#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;

#endregion

namespace PilotRevitAddin01
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        //public static Document doc = null;
        //public static UIApplication application = null;
        System.Windows.Forms.Form form = null;
        private IntPtr _revit_window; // 2019


        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;

            
            _revit_window = uiapp.MainWindowHandle; // 2019
            IWin32Window revit_window = new JtWindowHandle(uiapp.MainWindowHandle);


            try
            {
                form = new ListForm(doc);
                form.StartPosition = FormStartPosition.CenterParent;
                form.Show(revit_window);
                //form.ShowDialog(revit_window);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
