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
    internal class AppWithClasses : IExternalApplication
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
            ButtonClass data1 = new ButtonClass("Tool1", "Tool 1", "RAB_Session_06_Challenge.Command",
                RAB_Session_06_Challenge.Properties.Resources.Blue_32,
                RAB_Session_06_Challenge.Properties.Resources.Blue_16, "This is tool 1 tooltip");
            ButtonClass data2 = new ButtonClass("Tool2", "Tool 2", "RAB_Session_06_Challenge.Command",
                RAB_Session_06_Challenge.Properties.Resources.Blue_32,
                RAB_Session_06_Challenge.Properties.Resources.Blue_16, "This is tool 1 tooltip");
            ButtonClass data3 = new ButtonClass("Tool3", "Tool 3", "RAB_Session_06_Challenge.Command",
                RAB_Session_06_Challenge.Properties.Resources.Blue_32,
                RAB_Session_06_Challenge.Properties.Resources.Blue_16, "This is tool 1 tooltip");
            ButtonClass data4 = new ButtonClass("Tool4", "Tool 4", "RAB_Session_06_Challenge.Command",
                RAB_Session_06_Challenge.Properties.Resources.Blue_32,
                RAB_Session_06_Challenge.Properties.Resources.Blue_16, "This is tool 1 tooltip");
            ButtonClass data5 = new ButtonClass("Tool5", "Tool 5", "RAB_Session_06_Challenge.Command",
                RAB_Session_06_Challenge.Properties.Resources.Blue_32,
                RAB_Session_06_Challenge.Properties.Resources.Blue_16, "This is tool 1 tooltip");
            ButtonClass data6 = new ButtonClass("Tool6", "Tool 6", "RAB_Session_06_Challenge.Command",
                RAB_Session_06_Challenge.Properties.Resources.Blue_32,
                RAB_Session_06_Challenge.Properties.Resources.Blue_16, "This is tool 1 tooltip");
            ButtonClass data7 = new ButtonClass("Tool7", "Tool 7", "RAB_Session_06_Challenge.Command",
                RAB_Session_06_Challenge.Properties.Resources.Blue_32,
                RAB_Session_06_Challenge.Properties.Resources.Blue_16, "This is tool 1 tooltip");
            ButtonClass data8 = new ButtonClass("Tool8", "Tool 8", "RAB_Session_06_Challenge.Command",
                RAB_Session_06_Challenge.Properties.Resources.Blue_32,
                RAB_Session_06_Challenge.Properties.Resources.Blue_16, "This is tool 1 tooltip");
            ButtonClass data9 = new ButtonClass("Tool9", "Tool 9", "RAB_Session_06_Challenge.Command",
                RAB_Session_06_Challenge.Properties.Resources.Blue_32,
                RAB_Session_06_Challenge.Properties.Resources.Blue_16, "This is tool 1 tooltip");
            ButtonClass data10 = new ButtonClass("Tool10", "Tool 10", "RAB_Session_06_Challenge.Command",
                RAB_Session_06_Challenge.Properties.Resources.Blue_32,
                RAB_Session_06_Challenge.Properties.Resources.Blue_16, "This is tool 1 tooltip");

            SplitButtonData splitData1 = new SplitButtonData("Split1", "Split\rButton");
            PulldownButtonData pulldownData1 = new PulldownButtonData("Pulldown1", "Pulldown\rButton");

            // 4a. Add references to project
            // PresentationCore, PresentationFramework, WindowsBase, System.Drawing
            // add using statement for System.Windows.Media.Imaging

            pulldownData1.LargeImage = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Red_32);
            pulldownData1.Image = BitmapToImageSource(RAB_Session_06_Challenge.Properties.Resources.Red_16);

            // 6. Create buttons
            PushButton button1 = panel.AddItem(data1.Data) as PushButton;
            PushButton button2 = panel.AddItem(data2.Data) as PushButton;

            panel.AddStackedItems(data3.Data, data4.Data, data5.Data);

            SplitButton split = panel.AddItem(splitData1) as SplitButton;
            split.AddPushButton(data6.Data);
            split.AddPushButton(data7.Data);

            PulldownButton pulldown = panel.AddItem(pulldownData1) as PulldownButton;
            pulldown.AddPushButton(data8.Data);
            pulldown.AddPushButton(data9.Data);
            pulldown.AddPushButton(data10.Data);

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

    public class ButtonClass
    {
        public PushButtonData Data { get; set; }

        public ButtonClass(string name, string text, string className, System.Drawing.Bitmap largeImage,
            System.Drawing.Bitmap smallImage, string toolTip)
        {
            Data = new PushButtonData(name, text, GetAssemblyName(), className);
            Data.ToolTip = toolTip;
            Data.LargeImage = BitmapToImageSource(largeImage);
            Data.Image = BitmapToImageSource(smallImage);
        }

        private string GetAssemblyName()
        {
            return Assembly.GetExecutingAssembly().Location;
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
    }
}
