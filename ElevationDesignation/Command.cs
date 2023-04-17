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

            // get data from the form

            string curElev = curForm.GetComboBoxCurElevSelectedItem();
            string newElev = curForm.GetComboBoxNewElevSelectedItem();

            // set some variables

            string curFilter = "";

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

            string newFilter = "";

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
            
            // start the transaction

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Replace Elevation Designation");

                // loop through the views and replace the elevation designation

                foreach (View curView in viewsList)
                {
                    if (curView.Name.Contains(curElev + " "))
                        curView.Name = curView.Name.Replace(curElev + " ", newElev + " ");
                }

                // loop through the sheets

                foreach (ViewSheet curSheet in sheetsList)
                {

                    // set some variables

                    string grpName = GetParameterValueByName(curSheet, "Group");
                    string grpFilter = GetParameterValueByName(curSheet, "Code Filter");

                    // change elevation designation in sheet number

                    if (curSheet.SheetNumber.Contains(curElev.ToLower()))
                        curSheet.SheetNumber = curSheet.SheetNumber.Replace(curElev.ToLower(), newElev.ToLower());

                    // change the group name

                   string grpNewName = GetLastCharacterInString(grpName, curElev, newElev);

                    if (grpName.Contains(curElev))
                        SetParameterByName(curSheet, "Group", grpNewName);

                    // update the code filter

                    if (grpName.Contains(curElev) && grpFilter != null && grpFilter.Contains(curFilter))
                        SetParameterByName(curSheet, "Code Filter", newFilter);                               
                }

                // add code to replace schedules

                // get a list of the viewports on the cover sheet

                // commit the changes

                t.Commit();

                // alert the user

                TaskDialog.Show("Complete", "Changed Elevation " + curElev + " to Elevation " + newElev);

                return Result.Succeeded;
            }           
        }

        private string GetLastCharacterInString(string grpName, string curElev, string newElev)
        {
            char lastChar = grpName[grpName.Length - 1];
            

            string grpLastChar = lastChar.ToString();
            

            if (grpLastChar == curElev)
            {
                return "Elevation " + newElev;
            }
            else
            {
                return newElev + grpName.Substring(1);
            }
        }

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

        private string SetParameterByName(Element curElem, string paramName, string value)
        {
            Parameter curParam = GetParameterByName(curElem, paramName);

            curParam.Set(value);
            return curParam.ToString();
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

        public static ViewSheet GetSheetByName(Document curDoc, string sheetName)
        {
            //get all sheets
            List<ViewSheet> curSheets = GetAllSheets(curDoc);

            //loop through sheets and check sheet name
            foreach (ViewSheet curSheet in curSheets)
            {
                if (curSheet.Name == sheetName)
                {
                    return curSheet;
                }
            }

            return null;
        }

        public static List<Viewport> GetAllViewports(Document curDoc)
        {
            //get all viewports
            FilteredElementCollector vpCollector = new FilteredElementCollector(curDoc);
            vpCollector.OfCategory(BuiltInCategory.OST_Viewports);

            //output viewports to list
            List<Viewport> vpList = new List<Viewport>();
            foreach (Viewport curVP in vpCollector)
            {
                //add to list
                vpList.Add(curVP);
            }

            return vpList;
        }

        public static List<string> GetAllViewportTypes(Document curDoc)
        {
            List<string> returnList = new List<string>();
            List<Viewport> vpList = GetAllViewports(curDoc);

            foreach (Viewport vp in vpList)
            {
                returnList.Add(vp.Name);
            }

            return returnList;
        }

        public static List<Viewport> GetAllViewportsInView(Document curDoc, ElementId curViewID)
        {
            FilteredElementCollector vpCollector = new FilteredElementCollector(curDoc, curViewID);
            vpCollector.OfCategory(BuiltInCategory.OST_Viewports);

            //output viewports to list
            List<Viewport> vpList = new List<Viewport>();
            foreach (Viewport curVP in vpCollector)
            {
                //add to list
                vpList.Add(curVP);
            }

            return vpList;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
