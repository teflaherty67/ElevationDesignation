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
    public class cmdElevDesignationTest : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Document curDoc = uidoc.Document;

                // Open the form and handle DialogResult
                frmReplaceElevation curForm = OpenElevationReplacementForm();
                if (curForm.DialogResult == false)
                {
                    return Result.Failed;
                }

                // Get data from the form
                string curElev = curForm.GetComboBoxCurElevSelectedItem();
                string newElev = curForm.GetComboBoxNewElevSelectedItem();
                string codeMasonry = curForm.GetComboBoxCodeMasonrySelectedItem();

                // Create a filter for newElev
                string newFilter = CreateFilterForNewElevation(newElev);

                // Get all views and sheets
                List<View> viewsList = Utils.GetAllViews(curDoc);
                List<ViewSheet> sheetsList = Utils.GetAllSheets(curDoc);

                // Check if all schedules exist for newElev
                if (!CheckSchedulesExist(curDoc, curElev, newElev))
                {
                    ShowScheduleErrorDialog("The schedules for the new elevation do not exist or do not follow the proper naming convention. Please create the schedules or correct the schedule names and try again.");
                    return Result.Failed;
                }

                // Perform the elevation replacement
                if (ReplaceElevationDesignation(curDoc, curElev, newElev, codeMasonry, newFilter, viewsList, sheetsList, uidoc))
                {
                    ShowSuccessDialog($"Changed Elevation {curElev} to Elevation {newElev}");
                    return Result.Succeeded;
                }
                else
                {
                    return Result.Failed;
                }
            }
            catch (Exception ex)
            {
                ShowErrorDialog($"An error occurred: {ex.Message}");
                return Result.Failed;
            }
        }

        private frmReplaceElevation OpenElevationReplacementForm()
        {
            frmReplaceElevation curForm = new frmReplaceElevation()
            {
                Width = 340,
                Height = 200,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            curForm.ShowDialog();
            return curForm;
        }

        private string CreateFilterForNewElevation(string newElev)
        {
            switch (newElev)
            {
                case "A":
                    return "1";
                case "B":
                    return "2";
                case "C":
                    return "3";
                case "D":
                    return "4";
                case "S":
                    return "5";
                case "T":
                    return "6";
                default:
                    return "";
            }
        }

        private bool CheckSchedulesExist(Document curDoc, string curElev, string newElev)
        {
            List<ViewSchedule> curElevList = Utils.GetAllSchedulesByElevation(curDoc, curElev);
            List<ViewSchedule> newElevList = Utils.GetAllSchedulesByElevation(curDoc, newElev);
            return curElevList.Count != 0 && newElevList.Count == curElevList.Count;
        }

        private bool ReplaceElevationDesignation(Document curDoc, string curElev, string newElev, string codeMasonry, string newFilter, List<View> viewsList, List<ViewSheet> sheetsList, UIDocument uidoc)
        {
            using (TransactionGroup tGroup = new TransactionGroup(curDoc))
            {
                using (Transaction t = new Transaction(curDoc))
                {
                    tGroup.Start("Replace Elevation Designation");

                    // Rename views and sheets
                    if (!RenameViewsAndSheets(curDoc, curElev, newElev, viewsList, sheetsList))
                    {
                        return false;
                    }

                    // Replace Cover schedules
                    if (!ReplaceCoverSchedules(curDoc, curElev, newElev, uidoc))
                    {
                        return false;
                    }

                    // Replace Roof schedules
                    if (!ReplaceRoofSchedules(curDoc, curElev, newElev, uidoc))
                    {
                        return false;
                    }

                    t.Commit();
                    tGroup.Assimilate();
                }
            }
            return true;
        }

        private bool RenameViewsAndSheets(Document curDoc, string curElev, string newElev, List<View> viewsList, List<ViewSheet> sheetsList)
        {
            // Implement the renaming logic here
            // ...

            return true; // Return true if renaming was successful, false otherwise
        }

        private bool ReplaceCoverSchedules(Document curDoc, string curElev, string newElev, UIDocument uidoc)
        {
            // Implement the Cover schedule replacement logic here
            // ...

            return true; // Return true if replacement was successful, false otherwise
        }

        private bool ReplaceRoofSchedules(Document curDoc, string curElev, string newElev, UIDocument uidoc)
        {
            // Implement the Roof schedule replacement logic here
            // ...

            return true; // Return true if replacement was successful, false otherwise
        }

        private void ShowSuccessDialog(string message)
        {
            TaskDialog tdSuccess = new TaskDialog("Complete");
            tdSuccess.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
            tdSuccess.Title = "Complete";
            tdSuccess.TitleAutoPrefix = false;
            tdSuccess.MainContent = message;
            tdSuccess.CommonButtons = TaskDialogCommonButtons.Close;
            tdSuccess.Show();
        }

        private void ShowScheduleErrorDialog(string message)
        {
            TaskDialog tdCurSchedError = new TaskDialog("Error");
            tdCurSchedError.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
            tdCurSchedError.Title = "Replace Elevation Designation";
            tdCurSchedError.TitleAutoPrefix = false;
            tdCurSchedError.MainContent = message;
            tdCurSchedError.CommonButtons = TaskDialogCommonButtons.Close;
            tdCurSchedError.Show();
        }

        private void ShowErrorDialog(string message)
        {
            TaskDialog tdError = new TaskDialog("Error");
            tdError.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
            tdError.Title = "Error";
            tdError.TitleAutoPrefix = false;
            tdError.MainContent = message;
            tdError.CommonButtons = TaskDialogCommonButtons.Close;
            tdError.Show();
        }

        public static String GetMethod()
        {
            return MethodBase.GetCurrentMethod().DeclaringType?.FullName;
        }
    }
}

