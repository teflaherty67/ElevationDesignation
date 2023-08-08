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
                Width = 340,
                Height = 200,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            curForm.ShowDialog();

            if (curForm.DialogResult == false)
            {
                return Result.Failed;
            }

            // get data from the form
            string curElev = curForm.GetComboBoxCurElevSelectedItem();
            string newElev = curForm.GetComboBoxNewElevSelectedItem();
            string codeMasonry = curForm.GetComboBoxCodeMasonrySelectedItem();

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

            // get all the views & sheets
            List<View> viewsList = Utils.GetAllViews(curDoc);
            List<ViewSheet> sheetsList = Utils.GetAllSheets(curDoc);

            // check if all the schedules exist for newElev
            List<ViewSchedule> curElevList = Utils.GetAllSchedulesByElevation(curDoc, curElev);
            List<ViewSchedule> newElevList = Utils.GetAllSchedulesByElevation(curDoc, newElev);

            if (curElevList.Count != 0 && newElevList.Count == curElevList.Count)
            {
                // create a transaction group
                using (TransactionGroup tGroup = new TransactionGroup(curDoc))
                {
                    // create a transaction
                    using (Transaction t = new Transaction(curDoc))
                    {
                        // start the transaction group
                        tGroup.Start("Replace Elevation Designation");

                        #region Rename views & Sheets

                        // start the 1st transaction
                        t.Start("Rename views & sheets");

                        // create a counter for the views
                        int countView = 0;

                        // loop through the views
                        foreach (View curView in viewsList)
                        {
                            try
                            {
                                // rename the views
                                if (curView.Name.Contains(curElev + " "))
                                    curView.Name = curView.Name.Replace(curElev + " ", newElev + " ");
                                if (curView.Name.Contains(curElev + "-"))
                                    curView.Name = curView.Name.Replace(curElev + "-", newElev + " ");
                                if (curView.Name.Contains(curElev + "_"))
                                    curView.Name = curView.Name.Replace(curElev + "_", newElev + " ");
                            }
                            catch (Autodesk.Revit.Exceptions.ArgumentException)
                            {
                                // increment the counter
                                countView++;

                                continue;
                            }                            
                        }

                        // if the views already exist, alert the user & continue
                        if (countView > 0)
                        {
                            TaskDialog tdDupViews = new TaskDialog("Error");
                            tdDupViews.MainIcon = TaskDialogIcon.TaskDialogIconInformation;
                            tdDupViews.Title = "Duplicate View Names";
                            tdDupViews.TitleAutoPrefix = false;
                            tdDupViews.MainContent = "The views already exist";
                            tdDupViews.CommonButtons = TaskDialogCommonButtons.Close;

                            TaskDialogResult tdDupViewsRes = tdDupViews.Show();
                        }
                        
                        // create a counter for the sheets
                        int countSheets = 0;

                        // loop through the sheets & change sheet number & name
                        foreach (ViewSheet curSheet in sheetsList)
                        {
                            // set some varibales
                            string originalName = curSheet.Name;
                            string newName = "";

                            // remove the code filter from the sheet names
                            if (originalName.Length > 0 && originalName.Contains("-"))
                            {
                                string sheetName = originalName.Split('-')[0];

                                // check to see if the original name ends with "g"
                                if (originalName.EndsWith("g"))
                                {
                                    newName = sheetName + " -g";
                                }
                                else
                                {
                                    newName = sheetName;
                                }

                                // set the new sheet name
                                curSheet.Name = newName;
                            }

                            // rename the elevation sheets
                            if (curSheet.Name.StartsWith("Elevation"))
                                curSheet.Name = "Exterior Elevations";

                            try
                            {
                                // change elevation designation in sheet number
                                if (curSheet.SheetNumber.Contains(curElev.ToLower()))
                                    curSheet.SheetNumber = curSheet.SheetNumber.Replace(curElev.ToLower(), newElev.ToLower());
                            }
                            catch (Autodesk.Revit.Exceptions.ArgumentException)
                            {
                                // increment the counter
                                countSheets++;

                                continue;
                            }                            
                        }

                        // if the sheets already exist, alert the user & continue
                        if (countSheets > 0)
                        {
                            TaskDialog tdDupSheets = new TaskDialog("Error");
                            tdDupSheets.MainIcon = TaskDialogIcon.TaskDialogIconInformation;
                            tdDupSheets.Title = "Duplicate Sheet Names";
                            tdDupSheets.TitleAutoPrefix = false;
                            tdDupSheets.MainContent = "The sheets already exist";
                            tdDupSheets.CommonButtons = TaskDialogCommonButtons.Close;

                            TaskDialogResult tdDupSheetsRes = tdDupSheets.Show();

                            List<ViewSheet> newElevSheetList = Utils.GetAllSheetsByElevation(curDoc, newElev.ToLower());

                            foreach (ViewSheet curSheet in newElevSheetList)
                            {
                                // create some variables for parameters to be updated
                                Parameter curCat = Utils.GetParameterByNameAndWritable(curSheet, "Category");
                                string curGrp = Utils.GetParameterValueByName(curSheet, "Group");                               

                                if (curGrp.Contains(newElev))
                                {
                                    // update the category
                                    try
                                    {
                                        if (curCat.Definition.Name != "Active")
                                        {
                                            Utils.SetParameterByName(curSheet, "Category", "Active");
                                        }
                                        else
                                        {
                                            continue;
                                        }
                                    }
                                    catch (Exception)
                                    {
                                    }

                                    // update the masonry code
                                    Utils.SetParameterByName(curSheet, "Code Masonry", codeMasonry);

                                    // update the group name
                                    string[] curGroup = curGrp.Split('-', '|');

                                    string curCode = curGroup[1];

                                    string newCode = curGroup[0] + "-" + codeMasonry + "|" + curGroup[2] + "|" + curGroup[3] + "|" + curGroup[4];

                                    Utils.SetParameterByName(curSheet, "Group", newCode);
                                }                              
                            }
                        }

                        if (countSheets == 0)
                        {
                            List<ViewSheet> newSheetList = Utils.GetAllSheetsByElevation(curDoc, newElev.ToLower());

                            foreach (ViewSheet curSheet in newSheetList)
                            {
                                // create some variables for parameters to be updated
                                string curCat = Utils.GetParameterValueByName(curSheet, "Category");
                                string curGrp = Utils.GetParameterValueByName(curSheet, "Group");

                                string grpNewName = Utils.GetLastCharacterInString(curGrp, curElev, newElev);

                                // change the group name
                                if (curGrp.Contains(curElev))
                                    Utils.SetParameterByName(curSheet, "Group", grpNewName);

                                // update the elevation designation
                                if (curGrp.Contains(curElev))
                                    Utils.SetParameterByName(curSheet, "Elevation Designation", newElev);

                                // update the code filter
                                if (curGrp.Contains(curElev))
                                    Utils.SetParameterByName(curSheet, "Code Filter", newFilter);

                                // update the masonry code
                                if (curGrp.Contains(curElev))
                                    Utils.SetParameterByName(curSheet, "Code Masonry", codeMasonry);

                                // replace masonry code in group name
                                string[] curGroup = grpNewName.Split('-', '|');

                                string curCode = curGroup[1];

                                string newCode = curGroup[0] + "-" + codeMasonry + "|" + curGroup[2] + "|" + curGroup[3] + "|" + curGroup[4];

                                Utils.SetParameterByName(curSheet, "Group", newCode);
                            }                                
                        }               

                        // commit the 1st transaction
                        t.Commit();

                        #endregion

                        // set the cover for newElev as the active view
                        ViewSheet newCover;
                        newCover = Utils.GetSheetByElevationAndNameContains(curDoc, newElev, "Cover");

                        uidoc.ActiveView = newCover;

                        #region Replace Cover schedules

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

                        #endregion

                        // set the roof plan for newElev as the active view
                        ViewSheet newRoof;
                        newRoof = Utils.GetSheetByElevationAndNameContains(curDoc, newElev, "Roof Plan");

                        uidoc.ActiveView = newRoof;

                        #region Replace ventilation schedules

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

                        #endregion

                        tGroup.Assimilate();
                    }
                }

                #region Task Dialogs

                // alert the user
                TaskDialog tdSuccess = new TaskDialog("Complete");
                tdSuccess.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
                tdSuccess.Title = "Complete";
                tdSuccess.TitleAutoPrefix = false;
                tdSuccess.MainContent = "Changed Elevation " + curElev + " to Elevation " + newElev;
                tdSuccess.CommonButtons = TaskDialogCommonButtons.Close;

                TaskDialogResult tdSuccessRes = tdSuccess.Show();                

                return Result.Succeeded;
            }

            else if (curElevList.Count == 0)
            {
                TaskDialog tdCurSchedError = new TaskDialog("Error");
                tdCurSchedError.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
                tdCurSchedError.Title = "Replace Elevation Designation";
                tdCurSchedError.TitleAutoPrefix = false;
                tdCurSchedError.MainContent = "The schedules for the current elevation do not follow the proper naming convention; " +
                    "please correct the schedule names and try again.";
                tdCurSchedError.CommonButtons = TaskDialogCommonButtons.Close;

                TaskDialogResult tdCurErrorRes = tdCurSchedError.Show();
            }

            else if (newElevList.Count != curElevList.Count)
            {
                // if the schedules don't exist, alert the user & exit
                TaskDialog tdNewSchedError = new TaskDialog("Error");
                tdNewSchedError.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
                tdNewSchedError.Title = "Replace Elevation Designation";
                tdNewSchedError.TitleAutoPrefix = false;
                tdNewSchedError.MainContent = "The schedules for the new elevation do not exist, or do not follow the proper naming convention; " +
                    "please create the schedules, or correct the schedule names, and try again.";
                tdNewSchedError.CommonButtons = TaskDialogCommonButtons.Close;

                TaskDialogResult tdNewErrorRes = tdNewSchedError.Show();
            }
            return Result.Failed;
        }

        #endregion

        public static String GetMethod()
        {
            var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
            return method;
        }
    }
}
