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
using Forms = System.Windows;
using System.Windows.Input;
using System.Windows.Media;
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

            List<ViewSchedule> scheduleList = GetAllSchedulesByElevation(doc, newElev);

            // start the transaction

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Replace Elevation Designation");

                // check if the schedules for the new elevation exist

                if (scheduleList.Count != 0)
                {
                    // if yes, execute the command

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

                    // final step: replace the schedules on the sheets

                    // set the cover sheet as the actvie view

                    ViewSheet newSheet;
                    newSheet = GetSheetByElevationAndName(doc, newElev, "Cover");

                    uidoc.ActiveView = newSheet;

                    List<ScheduleSheetInstance> viewSchedules = GetAllScheduleSheetInstancesByNameAndView(doc, "Elevation " + curElev, uidoc.ActiveView);

                    foreach (ScheduleSheetInstance curSchedule in viewSchedules)
                    {
                        if (curSchedule.Name.Contains(curElev))
                        {
                            ElementId newSheetId = newSheet.Id;
                            string schedName = curSchedule.Name;
                            string newSchedName = schedName.Substring(0, schedName.Length - 2) + newElev;

                            ViewSchedule newSchedule = GetScheduleByName(doc, newSchedName); // equal to ID of schedule to replace existing

                            XYZ instanceLoc = curSchedule.Point;

                            doc.Delete(curSchedule.Id); // remove existing schedule

                            // add new schedule
                            ScheduleSheetInstance newSSI = ScheduleSheetInstance.Create(doc, newSheetId, newSchedule.Id, instanceLoc);
                        }
                    }

                    // commit the changes

                    t.Commit();

                    // alert the user

                    string msgSucceeded = "Changed Elevation " + curElev + " to Elevation " + newElev;
                    string titleSucceeded = "Complete";
                    Forms.MessageBoxButton btnSucceeded= Forms.MessageBoxButton.OK;

                    Forms.MessageBox.Show(msgSucceeded, titleSucceeded, btnSucceeded, Forms.MessageBoxImage.Information);

                    return Result.Succeeded;
                }

                else if (scheduleList.Count == 0)
                {
                    // if not, alert the user & exit

                    string msgFailed = "The schedules for the new elevation do not exist. Create the schedules and try again";
                    string titleFailed = "Warning";
                    Forms.MessageBoxButton btnFailed = Forms.MessageBoxButton.OK;
                    Forms.MessageBoxResult result = Forms.MessageBox.Show(msgFailed, titleFailed, btnFailed, Forms.MessageBoxImage.Warning);                                           
                }

                return Result.Failed;
            }
        }

        private List<ViewSchedule> GetAllSchedulesByElevation(Document doc, string newElev)
        {
            List<ViewSchedule> scheduleList = GetAllSchedules(doc);

            List<ViewSchedule> returnList = new List<ViewSchedule>();

            foreach (ViewSchedule curVS in scheduleList)
            {
                if (curVS.Name.Contains(newElev))
                {
                    returnList.Add(curVS);
                }

                return returnList;
            }

            return null;
        }

        private List<ViewSchedule> GetAllSchedules(Document doc)
        {
            List<ViewSchedule> schedList = new List<ViewSchedule>();

            FilteredElementCollector curCollector = new FilteredElementCollector(doc);
            curCollector.OfClass(typeof(ViewSchedule));

            //loop through views and check if schedule - if so then put into schedule list
            foreach (ViewSchedule curView in curCollector)
            {
                if (curView.ViewType == ViewType.Schedule)
                {
                    schedList.Add((ViewSchedule)curView);
                }
            }

            return schedList;
        }

        private ViewSchedule GetScheduleByName(Document doc, string v)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(ViewSchedule));

            foreach (ViewSchedule curSchedule in collector)
            {
                if (curSchedule.Name == v)
                    return curSchedule;
            }

            return null;
        }

        private List<ScheduleSheetInstance> GetAllScheduleSheetInstancesByNameAndView(Document doc, string elevName, View activeView)
        {
            List<ScheduleSheetInstance> ssiList = GetAllScheduleSheetInstancesByView(doc, activeView);

            List<ScheduleSheetInstance> returnList = new List<ScheduleSheetInstance>();

            foreach (ScheduleSheetInstance curInstance in ssiList)
            {
                if (curInstance.Name.Contains(elevName))
                    returnList.Add(curInstance);
            }

            return returnList;
        }

        private List<ScheduleSheetInstance> GetAllScheduleSheetInstancesByView(Document doc, View activeView)
        {
            FilteredElementCollector colSSI = new FilteredElementCollector(doc, activeView.Id);
            colSSI.OfClass(typeof(ScheduleSheetInstance));

            List<ScheduleSheetInstance> returnList = new List<ScheduleSheetInstance>();

            foreach (ScheduleSheetInstance curInstance in colSSI)
            {
                returnList.Add(curInstance);
            }

            return returnList;
        }

        private ViewSheet GetSheetByElevationAndName(Document doc, string newElev, string sheetName)
        {
            List<ViewSheet> sheetLIst = GetAllSheets(doc);

            foreach (ViewSheet curVS in sheetLIst)
            {
                if (curVS.SheetNumber.Contains(newElev) && curVS.Name == sheetName)
                {
                    return curVS;
                }
            }

            return null;
        }

        private List<ScheduleSheetInstance> GetAllScheduleSheetInstancesByName(Document doc, string elevName)
        {
            List<ScheduleSheetInstance> ssiList = GetAllScheduleSheetInstances(doc);

            List<ScheduleSheetInstance> returnList = new List<ScheduleSheetInstance>();

            foreach (ScheduleSheetInstance curInstance in ssiList)
            {
                if (curInstance.Name.Contains(elevName))
                    returnList.Add(curInstance);
            }

            return returnList;
        }

        private List<ScheduleSheetInstance> GetAllScheduleSheetInstances(Document doc)
        {
            FilteredElementCollector colSSI = new FilteredElementCollector(doc);
            colSSI.OfClass(typeof(ScheduleSheetInstance));

            List<ScheduleSheetInstance> returnList = new List<ScheduleSheetInstance>();

            foreach (ScheduleSheetInstance curInstance in colSSI)
            {
                returnList.Add(curInstance);
            }

            return returnList;
        }

        private List<ViewSheet> GetRoofSheet(Document doc)
        {
            List<ViewSheet> returnList = new List<ViewSheet>();

            // get all schedules
            List<ScheduleSheetInstance> scheduleSheetInstances = GetAllScheduleSheetInstances(doc);

            foreach (ScheduleSheetInstance curSSI in scheduleSheetInstances)
            {
                if (curSSI.Name.Contains("Roof"))
                {
                    ViewSheet curSheet = doc.GetElement(curSSI.OwnerViewId) as ViewSheet;
                    returnList.Add(curSheet);
                }
            }

            return returnList;
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

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}