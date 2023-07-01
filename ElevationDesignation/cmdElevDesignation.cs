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
    public class cmdElevDesignation : IExternalCommand
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

            // check if all the schedules exist for newElev
            List<ViewSchedule> curElevList = Utils.GetAllSchedulesByElevation(curDoc, curElev);
            List<ViewSchedule> newElevList = Utils.GetAllSchedulesByElevation(curDoc, newElev);

            if (newElevList.Count == curElevList.Count)
            {
                // create a transaction group
                using (TransactionGroup tGroup = new TransactionGroup(curDoc))
                {
                    // create a transaction
                    using (Transaction t = new Transaction(curDoc))
                    {
                        // start the transaction group
                        tGroup.Start("Replace Elevation Designation");

                        // start the 1st transaction
                        t.Start("Rename views & sheets");

                        // loop through the views
                        foreach (View curView in viewsList)
                        {
                            // rename the views
                            if (curView.Name.Contains(curElev + " "))
                                curView.Name = curView.Name.Replace(curElev + " ", newElev + " ");
                            if (curView.Name.Contains(curElev + "-"))
                                curView.Name = curView.Name.Replace(curElev + "-", newElev + " ");
                            if (curView.Name.Contains(curElev + "_"))
                                curView.Name = curView.Name.Replace(curElev + "_", newElev + " ");
                        }

                        // loop through the sheets
                        foreach (ViewSheet curSheet in sheetsList)
                        {
                            // set some variables
                            string grpName = Utils.GetParameterValueByName(curSheet, "Group");
                            string grpFilter = Utils.GetParameterValueByName(curSheet, "Code Filter");
                            string codeMasonry = Utils.GetParameterValueByName(curSheet, "Code Masonry");
                            
                            // change elevation designation in sheet number
                            if (curSheet.SheetNumber.Contains(curElev.ToLower()))
                                curSheet.SheetNumber = curSheet.SheetNumber.Replace(curElev.ToLower(), newElev.ToLower());

                            // change the group name
                            string grpNewName = Utils.GetLastCharacterInString(grpName, curElev, newElev);

                            if (grpName.Contains(curElev))
                                Utils.SetParameterByName(curSheet, "Group", grpNewName);

                            // update the code filter
                            if (grpName.Contains(curElev))
                                Utils.SetParameterByName(curSheet, "Code Filter", newFilter);

                            // update the masonry code
                            if (grpName.Contains(newElev))
                                Utils.SetParameterByName(curSheet, "Code Masonry", "0");

                            if (grpNewName.StartsWith(newElev))
                            {
                                string[] curGroup = grpNewName.Split('-', '|');

                                string curCode = curGroup[1];

                                string newCode = curGroup[0] + "-" + "0" + "|" + curGroup[2] + "|" + curGroup[3] + "|" + curGroup[4];

                                Utils.SetParameterByName(curSheet, "Group", newCode);
                            }
                        }

                        // commit the 1st transaction
                        t.Commit();

                        // set the cover for newElev as the active view
                        ViewSheet newCover;
                        newCover = Utils.GetSheetByElevationAndNameContains(curDoc, newElev, "Cover");

                        uidoc.ActiveView = newCover;

                        // start the 2nd transaction
                        t.Start("Replace the Cover schedules");

                        // get all SSI in the active view
                        List<ScheduleSheetInstance> schedCover = Utils.GetAllScheduleSheetInstancesByNameAndView
                            (curDoc, "Elevation " + curElev, uidoc.ActiveView);

                        // loop through the SSI
                        foreach (ScheduleSheetInstance curSchedule in schedCover)
                        {
                            if (curSchedule.Name.Contains(curElev))
                            {
                                // set some variables
                                ElementId newSheetId = newCover.Id;
                                string schedName = curSchedule.Name;
                                string newSchedName = schedName.Substring(0, schedName.Length - 1) + newElev;

                                // get the schedule name
                                ViewSchedule newSchedule = Utils.GetScheduleByName(curDoc, newSchedName); // equal to ID of schedule to replace existing

                                // get the schedule location
                                XYZ instanceLoc = curSchedule.Point;

                                // remove the curElev schedule
                                curDoc.Delete(curSchedule.Id);

                                // add new schedule
                                ScheduleSheetInstance newSSI = ScheduleSheetInstance.Create(curDoc, newSheetId, newSchedule.Id, instanceLoc);
                            }
                        }

                        // commit the 2nd transaction
                        t.Commit();

                        // set the roof plan for newElev as the active view
                        ViewSheet newRoof;
                        newRoof = Utils.GetSheetByElevationAndNameContains(curDoc, newElev, "Roof");

                        uidoc.ActiveView = newRoof;

                        // start the 3rd transaction
                        t.Start("Replace the roof schedules");

                        // get all SSI in the active view
                        List<ScheduleSheetInstance> schedRoof = Utils.GetAllScheduleSheetInstancesByNameAndView
                            (curDoc, "Elevation " + curElev, uidoc.ActiveView);

                        // loop through the SSI
                        foreach (ScheduleSheetInstance curSchedule in schedRoof)
                        {
                            if (curSchedule.Name.Contains(curElev))
                            {
                                // set some variables
                                ElementId newSheetId = newRoof.Id;
                                string schedName = curSchedule.Name;
                                string newSchedName = schedName.Substring(0, schedName.Length - 1) + newElev;

                                // get the schedule name
                                ViewSchedule newSchedule = Utils.GetScheduleByName(curDoc, newSchedName); // equal to ID of schedule to replace existing

                                // get the schedule location
                                XYZ instanceLoc = curSchedule.Point;

                                // remove the curElev schedule
                                curDoc.Delete(curSchedule.Id);

                                // add new schedule
                                ScheduleSheetInstance newSSI = ScheduleSheetInstance.Create(curDoc, newSheetId, newSchedule.Id, instanceLoc);
                            }
                        }

                        // commit the 3rd transaction
                        t.Commit();

                        tGroup.Assimilate();
                    }
                }

                // alert the user
                string msgSucceeded = "Changed Elevation " + curElev + " to Elevation " + newElev;
                string titleSucceeded = "Complete";
                Forms.MessageBoxButton btnSucceeded = Forms.MessageBoxButton.OK;

                Forms.MessageBox.Show(msgSucceeded, titleSucceeded, btnSucceeded, Forms.MessageBoxImage.Information);

                return Result.Succeeded;
            }

            else if (newElevList.Count != curElevList.Count)
            {
                // if the schedules don't exist, alert the user & exit
                string msgFailed = "The schedules for the new elevation do not exist. Create the schedules and try again";
                string titleFailed = "Warning";
                Forms.MessageBoxButton btnFailed = Forms.MessageBoxButton.OK;
                Forms.MessageBoxResult result = Forms.MessageBox.Show(msgFailed, titleFailed, btnFailed, Forms.MessageBoxImage.Warning);
            }
            return Result.Failed;
        }

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
