#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Media.Imaging;

#endregion

namespace RAB_Session_06_Challenge
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            string assemblyName = GetAssemblyName();

            // 1. Create ribbon tab
            try
            {
                app.CreateRibbonTab("Revit Add-in Bootcamp");
            }
            catch (Exception)
            {
                Debug.Print("Tab already exists.");
            }

            // 2. Create ribbon panel 
            //RibbonPanel panel = app.CreateRibbonPanel("Revit Add-in Bootcamp", "Revit Tools");
            RibbonPanel panel = CreateRibbonPanel(app, "Revit Add-in Bootcamp", "Revit Tools");

            // 3. Create button data instances
            PushButtonData data1 = new PushButtonData("Tool1", "Tool 1", assemblyName, "RAB_Session_06_Challenge.Command");
            PushButtonData data2 = new PushButtonData("Tool2", "Tool 2", assemblyName, "RAB_Session_06_Challenge.Command");
            PushButtonData data3 = new PushButtonData("Tool3", "Tool 3", assemblyName, "RAB_Session_06_Challenge.Command");
            PushButtonData data4 = new PushButtonData("Tool4", "Tool 4", assemblyName, "RAB_Session_06_Challenge.Command");
            PushButtonData data5 = new PushButtonData("Tool5", "Tool 5", assemblyName, "RAB_Session_06_Challenge.Command");
            PushButtonData data6 = new PushButtonData("Tool6", "Tool 6", assemblyName, "RAB_Session_06_Challenge.Command");
            PushButtonData data7 = new PushButtonData("Tool7", "Tool 7", assemblyName, "RAB_Session_06_Challenge.Command");
            PushButtonData data8 = new PushButtonData("Tool8", "Tool 8", assemblyName, "RAB_Session_06_Challenge.Command");
            PushButtonData data9 = new PushButtonData("Tool9", "Tool 9", assemblyName, "RAB_Session_06_Challenge.Command");
            PushButtonData data10 = new PushButtonData("Tool10", "Tool 10", assemblyName, "RAB_Session_06_Challenge.Command");

            SplitButtonData splitData1 = new SplitButtonData("Split1", "Split\rButton");
            PulldownButtonData pulldownData1 = new PulldownButtonData("Pulldown1", "Pulldown\rButton");

            // 4a. Add references to project
            // PresentationCore, PresentationFramework, WindowsBase, System.Drawing
            // add using statement for System.Windows.Media.Imaging

            // 4. Add images to button data

            return Result.Succeeded;
        }

        private string GetAssemblyName()
        {
            return Assembly.GetExecutingAssembly().Location;
        }

        private RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel currentPanel = GetRibbonPanelByName(app, tabName, panelName);

            if (currentPanel == null)
                currentPanel = app.CreateRibbonPanel(tabName, panelName);

            return currentPanel;
        }

        private RibbonPanel GetRibbonPanelByName(UIControlledApplication app, string tabName, string panelName)
        {
            foreach(RibbonPanel tmpPanel in app.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                    return tmpPanel;
            }

            return null;
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }
    }
}
