#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;

#endregion

namespace ElevationDesignation
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;

            // put any code needed for the form here

            // open form
            frmReplaceElevation curForm = new frmReplaceElevation()
            {
                Width = 320,
                Height = 180,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            curForm.ShowDialog();

            // get form data and do something

            string curElev = curForm.GetComboBoxCurElevSelectedItem();
            string newElev = curForm.GetComboBoxNewElevSelectedItem();

            List<View> viewsList = GetAllViews(doc);

            List<ViewSheet> sheetsList = GetAllSheets(doc);

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Replace Elevation Designation");

                foreach (View curView in viewsList)
                {
                    if (curView.Name.Contains(curElev + " "))
                        curView.Name = curView.Name.Replace(curElev + " ", newElev + " ");
                }

                foreach (ViewSheet curSheet in sheetsList)
                {
                    if (curSheet.SheetNumber.Contains(curElev.ToLower()))
                        curSheet.SheetNumber = curSheet.SheetNumber.Replace(curElev.ToLower(), newElev.ToLower());
                }

                t.Commit();

                return Result.Succeeded;
            }           
        }

        private List<ViewSheet> GetAllSheets(Document doc)
        {
            throw new NotImplementedException();
        }

        public static List<View> GetAllViews(Document curDoc)
        {
            FilteredElementCollector colviews = new FilteredElementCollector(curDoc);
            colviews.OfCategory(BuiltInCategory.OST_Views);

            List<View> returnViews = new List<View>();
            foreach (View curView in colviews.ToElements())
            {
                returnViews.Add(curView);
            }

            return returnViews;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
