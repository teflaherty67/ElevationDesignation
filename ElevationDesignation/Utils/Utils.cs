using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace ElevationDesignation
{
    internal static class Utils
    {

        #region Parameters

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

        internal static string SetParameterByName(Element curElem, string paramName, string value)
        {
            Parameter curParam = GetParameterByNameAndWritable(curElem, paramName);

            curParam.Set(value);
            return curParam.ToString();
        }

        internal static Parameter GetParameterByName(Element curElem, string paramName)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name.ToString() == paramName)
                    return curParam;
            }

            return null;
        }

        internal static Parameter GetParameterByNameAndWritable(Element curElem, string paramName)
        {
            foreach (Parameter curParam in curElem.Parameters)
            {
                if (curParam.Definition.Name.ToString() == paramName && curParam.IsReadOnly == false)
                    return curParam;
            }

            return null;
        }

        #endregion

        #region Ribbon
        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel currentPanel = GetRibbonPanelByName(app, tabName, panelName);

            if (currentPanel == null)
                currentPanel = app.CreateRibbonPanel(tabName, panelName);

            return currentPanel;
        }

        internal static RibbonPanel GetRibbonPanelByName(UIControlledApplication app, string tabName, string panelName)
        {
            foreach (RibbonPanel tmpPanel in app.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                    return tmpPanel;
            }

            return null;
        }

        #endregion

        #region Schedules

        internal static List<ViewSchedule> GetAllSchedulesByElevation(Document doc, string newElev)
        {
            List<ViewSchedule> m_scheduleList = GetAllSchedules(doc);

            List<ViewSchedule> m_returnList = new List<ViewSchedule>();

            foreach (ViewSchedule curVS in m_scheduleList)
            {
                if (curVS.Name.EndsWith(newElev))
                {
                    m_returnList.Add(curVS);
                }                
            }

            return m_returnList;
        }

        internal static List<ViewSchedule> GetAllSchedules(Document doc)
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

        internal static ViewSchedule GetScheduleByName(Document doc, string v)
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

        internal static List<ScheduleSheetInstance> GetAllScheduleSheetInstancesByNameAndView(Document doc, string elevName, View activeView)
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

        internal static List<ScheduleSheetInstance> GetAllScheduleSheetInstancesByView(Document doc, View activeView)
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

        internal static List<ScheduleSheetInstance> GetAllScheduleSheetInstancesByName(Document doc, string elevName)
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

        internal static List<ScheduleSheetInstance> GetAllScheduleSheetInstances(Document doc)
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

        #endregion

        #region Sheets

        internal static List<ViewSheet> GetAllSheets(Document curDoc)
        {
            //get all sheets
            FilteredElementCollector m_colSheets = new FilteredElementCollector(curDoc);
            m_colSheets.OfCategory(BuiltInCategory.OST_Sheets);

            List<ViewSheet> m_returnSheets = new List<ViewSheet>();
            foreach (ViewSheet curSheet in m_colSheets.ToElements())
            {
                m_returnSheets.Add(curSheet);
            }

            return m_returnSheets;
        }

        internal static ViewSheet GetSheetByElevationAndNameContains(Document curDoc, string newElev, string sheetName)
        {
            List<ViewSheet> sheetList = GetAllSheets(curDoc);           

            foreach (ViewSheet curVS in sheetList)
            {
                if (curVS.SheetNumber.Contains(newElev.ToLower()) && curVS.Name.Contains(sheetName))
                {
                    return curVS;
                }
            }

            return null;
        }

        internal static List<ViewSheet> GetAllSheetsByElevation(Document curDoc, string elevDesignation)
        {
            //get all sheets
            List<ViewSheet> m_sheetList = GetAllSheets(curDoc);

            List<ViewSheet> m_returnSheets = new List<ViewSheet>();

            foreach (ViewSheet curSheet in m_sheetList)
            {
                if (curSheet.SheetNumber.Contains(elevDesignation))
                m_returnSheets.Add(curSheet);
            }

            return m_returnSheets;
        }

        #endregion

        #region Strings

        internal static string GetLastCharacterInString(string grpName, string curElev, string newElev)
        {
            char lastChar = grpName[grpName.Length - 1];

            char firstChar = grpName[0];

            string grpLastChar = lastChar.ToString();

            string grpFirstChar = firstChar.ToString();

            if (grpLastChar == curElev)
            {
                return "Elevation " + newElev;
            }
            else if (grpFirstChar == curElev)
            {
                return newElev + grpName.Substring(1);
            }
            else
            {
                return grpName;
            }
        }

        #endregion

        #region Views

        public static List<View> GetAllViews(Document curDoc)
        {
            FilteredElementCollector m_colviews = new FilteredElementCollector(curDoc);
            m_colviews.OfCategory(BuiltInCategory.OST_Views);

            List<View> m_returnViews = new List<View>();
            foreach (View curView in m_colviews.ToElements())
            {
                m_returnViews.Add(curView);
            }

            return m_returnViews;
        }

        #endregion

        #region Viewports

        internal static List<Viewport> GetAllViewports(Document curDoc)
        {
            //get all viewports
            FilteredElementCollector m_vpCollector = new FilteredElementCollector(curDoc);
            m_vpCollector.OfCategory(BuiltInCategory.OST_Viewports);

            //output viewports to list
            List<Viewport> m_vpList = new List<Viewport>();
            foreach (Viewport curVP in m_vpCollector)
            {
                //add to list
                m_vpList.Add(curVP);
            }

            return m_vpList;
        }     

        #endregion
    }
}
