#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
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
            PushButtonData data1 = new PushButtonData("Tool1", "Tool 1", assemblyName, "SmartPack_2022.AdskCmdAlign");
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
            data1.LargeImage = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Blue_32);
            data1.Image = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Blue_16);
            data2.LargeImage = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Red_32);
            data2.Image = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Red_16);
            data3.LargeImage = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Green_32);
            data3.Image = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Green_16);
            data4.LargeImage = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Red_32);
            data4.Image = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Red_16);
            data5.LargeImage = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Blue_32);
            data5.Image = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Blue_16);
            data6.LargeImage = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Blue_32);
            data6.Image = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Blue_16);
            data7.LargeImage = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Red_32);
            data7.Image = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Red_16);
            data8.LargeImage = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Green_32);
            data8.Image = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Green_16);
            data9.LargeImage = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Red_32);
            data9.Image = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Red_16);
            data10.LargeImage = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Blue_32);
            data10.Image = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Blue_16);

            pulldownData1.LargeImage = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Red_32);
            pulldownData1.Image = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Red_16);

            // 5. Add tooltips
            data1.ToolTip = "Tool 1 tool tip";
            data2.ToolTip = "Tool 2 tool tip";
            data3.ToolTip = "Tool 3 tool tip";
            data4.ToolTip = "Tool 4 tool tip";
            data5.ToolTip = "Tool 5 tool tip";
            data6.ToolTip = "Tool 6 tool tip";
            data7.ToolTip = "Tool 7 tool tip";
            data8.ToolTip = "Tool 8 tool tip";
            data9.ToolTip = "Tool 9 tool tip";
            data10.ToolTip = "Tool 10 tool tip";

            // 6. Create buttons
            PushButton button1 = panel.AddItem(data1) as PushButton;
            PushButton button2 = panel.AddItem(data2) as PushButton;

            panel.AddStackedItems(data3, data4, data5);

            SplitButton split = panel.AddItem(splitData1) as SplitButton;
            split.AddPushButton(data6);
            split.AddPushButton(data7);

            PulldownButton pulldown = panel.AddItem(pulldownData1) as PulldownButton;
            pulldown.AddPushButton(data8);
            pulldown.AddPushButton(data9);
            pulldown.AddPushButton(data10);

            // 7. Edit .addin file

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
        private BitmapImage BitmapToImageSource(System.Drawing.Bitmap bm)
        {
            using (MemoryStream mem = new MemoryStream())
            {
                bm.Save(mem, System.Drawing.Imaging.ImageFormat.Png);
                mem.Position = 0;
                BitmapImage bmi = new BitmapImage();
                bmi.BeginInit();
                bmi.StreamSource = mem;
                bmi.CacheOption = BitmapCacheOption.OnLoad;
                bmi.EndInit();

                return bmi;
            }
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }
    }
}
