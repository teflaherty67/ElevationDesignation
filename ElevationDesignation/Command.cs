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

            string curFilter;

            if (curElev == "A")
                curFilter = "1";
            else if (curElev == "B")
                curFilter = "2";
            else if (curElev == "C")
                curFilter = "3";
            else if (curElev == "D")
                curFilter = "4";
            else if (curElev == "S")
                curFilter = "5";
            else if (curElev == "T")
                curFilter = "6";

            string newFilter;

            if (newElev == "A")
                newFilter = "1";
            else if (newElev == "B")
                newFilter = "2";
            else if (newElev == "C")
                newFilter = "3";
            else if (newElev == "D")
                newFilter = "4";
            else if (newElev == "S")
                newFilter = "5";
            else if (newElev == "T")
                newFilter = "6";


            List<View> viewsList = GetAllViews(doc);

            List<ViewSheet> sheetsList = GetAllSheets(doc);

            List<ViewSheet> sheetGroup = GetAllSheetsByGroup(doc);

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
                    string grpName = GetParameterValueByName(curSheet, "Group");
                    string grpFilter = GetParameterValueByName(curSheet, "Code Filter");

                    if (curSheet.SheetNumber.Contains(curElev.ToLower()))
                        curSheet.SheetNumber = curSheet.SheetNumber.Replace(curElev.ToLower(), newElev.ToLower());

                    if (grpName.Contains(curElev))
                        grpName = SetParameterByName(curSheet, grpName, newElev);

                    if (grpFilter.Contains(curFilter))
                        grpFilter = SetParameterByName(curSheet, grpFilter, newFilter);
                }

                t.Commit();

                return Result.Succeeded;
            }           
        }

        //public static List<ViewSheet> GetAllSheetsByGroup(Document doc)
        //{
        //    List<ViewSheet> sheets = GetAllSheets(doc);

        //    List<ViewSheet> sheetGroup = new List<ViewSheet>();

        //    foreach (ViewSheet sheet in sheets)
        //    {
        //        if(sheet.GroupId = true)
        //        {
        //            sheetGroup.Add(sheet);
        //        }

        //        return sheetGroup;
        //    }

        //    return null;
        //}

        public static List<ViewSheet> GetAllSheets(Document curDoc)
        {
            //get all sheets
            FilteredElementCollector colSheets = new FilteredElementCollector(curDoc);
            colSheets.OfCategory(BuiltInCategory.OST_Sheets);

            List<ViewSheet> returnSheets = new List<ViewSheet>();
            foreach (ViewSheet curSheet in colSheets.ToElements())
            {
                returnSheets.Add(curSheet);
            }

            return returnSheets;
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

        internal static string GetParameterValueByName(Element element, string paramName)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);

            if (paramList != null)
                try
                {
                    Parameter param = paramList[0];
                    string paramValue = param.AsValueString();
                    return paramValue;
                }
                catch (System.ArgumentOutOfRangeException)
                {
                    return null;
                }

            return "";
        }

        public static Parameter GetParameterByName(Element curElem, string paramName)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name.ToString() == paramName)
                    return curParam;
            }

            return null;
        }

        internal static void SetParameterByName(Element element, string paramName, string value)
        {
            IList<Parameter> paramList = element.GetParameters(paramName);

            if (paramList != null)
            {
                Parameter param = paramList[0];

                param.Set(value);
            }
        }

        public static bool SetParameterValue(Element curElem, string paramName, string value)
        {
            Parameter curParam = GetParameterByName(curElem, paramName);

            if (curParam != null)
            {
                curParam.Set(value);
                return true;
            }

            return false;

        }

        public static bool SetParameterValue(Element curElem, string paramName, double value)
        {
            Parameter curParam = GetParameterByName(curElem, paramName);

            if (curParam != null)
            {
                curParam.Set(value);
                return true;
            }

            return false;

        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
