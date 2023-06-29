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
            Document curDoc = uidoc.Document;

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

            List<View> viewsList = Utils.GetAllViews(curDoc);

            List<ViewSheet> sheetsList = Utils.GetAllSheets(curDoc);

            List<ViewSchedule> scheduleList = Utils.GetAllSchedulesByElevation(curDoc, newElev);

            // start the transaction

            using (Transaction t = new Transaction(curDoc))
            {
                t.Start("Replace Elevation Designation");

                SubTransaction st1 = new SubTransaction(curDoc);

                st1.Start();

                // check if the schedules for the new elevation exist

                if (scheduleList.Count != 0)
                {
                    // if yes, execute the command

                    // loop through the views and replace the elevation designation

                    foreach (View curView in viewsList)
                    {
                        if (curView.Name.Contains(curElev + " ")
                            || curView.Name.Contains(curElev + "-")
                            || curView.Name.Contains(curElev + "_"))
                            curView.Name = curView.Name.Replace(curElev, newElev);
                    }

                    // loop through the sheets

                    foreach (ViewSheet curSheet in sheetsList)
                    {
                        // set some variables

                        string grpName = Utils.GetParameterValueByName(curSheet, "Group");
                        string grpFilter = Utils.GetParameterValueByName(curSheet, "Code Filter");

                        // change elevation designation in sheet number

                        if (curSheet.SheetNumber.Contains(curElev.ToLower()))
                            curSheet.SheetNumber = curSheet.SheetNumber.Replace(curElev.ToLower(), newElev.ToLower());

                        // change the group name

                        string grpNewName = Utils.GetLastCharacterInString(grpName, curElev, newElev);

                        if (grpName.Contains(curElev))
                            Utils.SetParameterByName(curSheet, "Group", grpNewName);

                        // update the code filter

                        if (grpName.Contains(curElev) && grpFilter != null && grpFilter.Contains(curFilter))
                            Utils.SetParameterByName(curSheet, "Code Filter", newFilter);
                    }

                    st1.Commit();

                    // final step: replace the schedules on the sheets

                    SubTransaction st2 = new SubTransaction(curDoc);

                    st2.Start();

                    // set the cover sheet as the actvie view

                    ViewSheet newSheet;
                    newSheet = Utils.GetSheetByElevationAndName(curDoc, newElev, "Cover");

                    uidoc.ActiveView = newSheet;

                    List<ScheduleSheetInstance> viewSchedules = Utils.GetAllScheduleSheetInstancesByNameAndView(curDoc, "Elevation " + curElev, uidoc.ActiveView);

                    foreach (ScheduleSheetInstance curSchedule in viewSchedules)
                    {
                        if (curSchedule.Name.Contains(curElev))
                        {
                            ElementId newSheetId = newSheet.Id;
                            string schedName = curSchedule.Name;
                            string newSchedName = schedName.Substring(0, schedName.Length - 2) + newElev;

                            ViewSchedule newSchedule = Utils.GetScheduleByName(curDoc, newSchedName); // equal to ID of schedule to replace existing

                            XYZ instanceLoc = curSchedule.Point;

                            curDoc.Delete(curSchedule.Id); // remove existing schedule

                            // add new schedule
                            ScheduleSheetInstance newSSI = ScheduleSheetInstance.Create(curDoc, newSheetId, newSchedule.Id, instanceLoc);
                        }
                    }

                    st2.Commit();

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

        

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}