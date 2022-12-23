#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Interop;
using System.Windows;
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
            ButtonClass data1 = new ButtonClass("Tool1", "Tool 1", "SmartPack_Shared.AdskCmdAlign",
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
        string blue32 = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAFSmlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4KPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iWE1QIENvcmUgNS41LjAiPgogPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4KICA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIgogICAgeG1sbnM6ZGM9Imh0dHA6Ly9wdXJsLm9yZy9kYy9lbGVtZW50cy8xLjEvIgogICAgeG1sbnM6ZXhpZj0iaHR0cDovL25zLmFkb2JlLmNvbS9leGlmLzEuMC8iCiAgICB4bWxuczp0aWZmPSJodHRwOi8vbnMuYWRvYmUuY29tL3RpZmYvMS4wLyIKICAgIHhtbG5zOnBob3Rvc2hvcD0iaHR0cDovL25zLmFkb2JlLmNvbS9waG90b3Nob3AvMS4wLyIKICAgIHhtbG5zOnhtcD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyIKICAgIHhtbG5zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIgogICAgeG1sbnM6c3RFdnQ9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZUV2ZW50IyIKICAgZXhpZjpQaXhlbFhEaW1lbnNpb249IjMyIgogICBleGlmOlBpeGVsWURpbWVuc2lvbj0iMzIiCiAgIGV4aWY6Q29sb3JTcGFjZT0iMSIKICAgdGlmZjpJbWFnZVdpZHRoPSIzMiIKICAgdGlmZjpJbWFnZUxlbmd0aD0iMzIiCiAgIHRpZmY6UmVzb2x1dGlvblVuaXQ9IjIiCiAgIHRpZmY6WFJlc29sdXRpb249Ijk2LjAiCiAgIHRpZmY6WVJlc29sdXRpb249Ijk2LjAiCiAgIHBob3Rvc2hvcDpDb2xvck1vZGU9IjMiCiAgIHBob3Rvc2hvcDpJQ0NQcm9maWxlPSJzUkdCIElFQzYxOTY2LTIuMSIKICAgeG1wOk1vZGlmeURhdGU9IjIwMjItMDgtMDNUMTE6MDA6NDUtMDQ6MDAiCiAgIHhtcDpNZXRhZGF0YURhdGU9IjIwMjItMDgtMDNUMTE6MDA6NDUtMDQ6MDAiPgogICA8ZGM6dGl0bGU+CiAgICA8cmRmOkFsdD4KICAgICA8cmRmOmxpIHhtbDpsYW5nPSJ4LWRlZmF1bHQiPkljb25zPC9yZGY6bGk+CiAgICA8L3JkZjpBbHQ+CiAgIDwvZGM6dGl0bGU+CiAgIDx4bXBNTTpIaXN0b3J5PgogICAgPHJkZjpTZXE+CiAgICAgPHJkZjpsaQogICAgICBzdEV2dDphY3Rpb249InByb2R1Y2VkIgogICAgICBzdEV2dDpzb2Z0d2FyZUFnZW50PSJBZmZpbml0eSBQaG90byAxLjkuMiIKICAgICAgc3RFdnQ6d2hlbj0iMjAyMi0wOC0wM1QxMTowMDo0NS0wNDowMCIvPgogICAgPC9yZGY6U2VxPgogICA8L3htcE1NOkhpc3Rvcnk+CiAgPC9yZGY6RGVzY3JpcHRpb24+CiA8L3JkZjpSREY+CjwveDp4bXBtZXRhPgo8P3hwYWNrZXQgZW5kPSJyIj8+pNjWXwAAAYFpQ0NQc1JHQiBJRUM2MTk2Ni0yLjEAACiRdZG7SwNBEIe/JIriG7SwsAiiqRKJClEbiwSNglokEYzaJJeXkMdxF5FgK9gGFEQbX4X+BdoK1oKgKIJYirWijYZzLhEiYmaZnW9/uzPszoI1lFYyep0bMtm8FvB77QvhRXvDM43YaGMER0TR1dngZIia9nGHxYw3LrNW7XP/WnMsritgaRQeV1QtLzwlPLOWV03eFu5SUpGY8KmwU5MLCt+aerTCLyYnK/xlshYK+MDaIWxP/uLoL1ZSWkZYXk5fJr2q/NzHfElLPDsflNgr3oNOAD9e7EwzgQ8Pg4zJ7MHFEAOyoka+u5w/R05yFZlVCmiskCRFHqeoq1I9LjEhelxGmoLZ/7991RPDQ5XqLV6ofzKMt35o2IJS0TA+Dw2jdAS2R7jIVvNzBzD6LnqxqvXtQ/sGnF1WtegOnG9C94Ma0SJlySZuTSTg9QRaw9B5DU1LlZ797HN8D6F1+aor2N0Dh5xvX/4Gl2ln/HOw8/IAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAA/SURBVEiJ7dahDQAwEEJRaLr/ytRdcgNgmo9C8SxOomqqQJI7vbBvSaewuwIAAAAAAAAAAPALMN/UJcDt+/4AUBENNtB9IZ4AAAAASUVORK5CYII=";
        string green32 = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAFSmlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4KPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iWE1QIENvcmUgNS41LjAiPgogPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4KICA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIgogICAgeG1sbnM6ZGM9Imh0dHA6Ly9wdXJsLm9yZy9kYy9lbGVtZW50cy8xLjEvIgogICAgeG1sbnM6ZXhpZj0iaHR0cDovL25zLmFkb2JlLmNvbS9leGlmLzEuMC8iCiAgICB4bWxuczp0aWZmPSJodHRwOi8vbnMuYWRvYmUuY29tL3RpZmYvMS4wLyIKICAgIHhtbG5zOnBob3Rvc2hvcD0iaHR0cDovL25zLmFkb2JlLmNvbS9waG90b3Nob3AvMS4wLyIKICAgIHhtbG5zOnhtcD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyIKICAgIHhtbG5zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIgogICAgeG1sbnM6c3RFdnQ9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZUV2ZW50IyIKICAgZXhpZjpQaXhlbFhEaW1lbnNpb249IjMyIgogICBleGlmOlBpeGVsWURpbWVuc2lvbj0iMzIiCiAgIGV4aWY6Q29sb3JTcGFjZT0iMSIKICAgdGlmZjpJbWFnZVdpZHRoPSIzMiIKICAgdGlmZjpJbWFnZUxlbmd0aD0iMzIiCiAgIHRpZmY6UmVzb2x1dGlvblVuaXQ9IjIiCiAgIHRpZmY6WFJlc29sdXRpb249Ijk2LjAiCiAgIHRpZmY6WVJlc29sdXRpb249Ijk2LjAiCiAgIHBob3Rvc2hvcDpDb2xvck1vZGU9IjMiCiAgIHBob3Rvc2hvcDpJQ0NQcm9maWxlPSJzUkdCIElFQzYxOTY2LTIuMSIKICAgeG1wOk1vZGlmeURhdGU9IjIwMjItMDgtMDNUMTE6MDE6MTUtMDQ6MDAiCiAgIHhtcDpNZXRhZGF0YURhdGU9IjIwMjItMDgtMDNUMTE6MDE6MTUtMDQ6MDAiPgogICA8ZGM6dGl0bGU+CiAgICA8cmRmOkFsdD4KICAgICA8cmRmOmxpIHhtbDpsYW5nPSJ4LWRlZmF1bHQiPkljb25zPC9yZGY6bGk+CiAgICA8L3JkZjpBbHQ+CiAgIDwvZGM6dGl0bGU+CiAgIDx4bXBNTTpIaXN0b3J5PgogICAgPHJkZjpTZXE+CiAgICAgPHJkZjpsaQogICAgICBzdEV2dDphY3Rpb249InByb2R1Y2VkIgogICAgICBzdEV2dDpzb2Z0d2FyZUFnZW50PSJBZmZpbml0eSBQaG90byAxLjkuMiIKICAgICAgc3RFdnQ6d2hlbj0iMjAyMi0wOC0wM1QxMTowMToxNS0wNDowMCIvPgogICAgPC9yZGY6U2VxPgogICA8L3htcE1NOkhpc3Rvcnk+CiAgPC9yZGY6RGVzY3JpcHRpb24+CiA8L3JkZjpSREY+CjwveDp4bXBtZXRhPgo8P3hwYWNrZXQgZW5kPSJyIj8+8OGIawAAAYFpQ0NQc1JHQiBJRUM2MTk2Ni0yLjEAACiRdZG7SwNBEIe/JIriG7SwsAiiqRKJClEbiwSNglokEYzaJJeXkMdxF5FgK9gGFEQbX4X+BdoK1oKgKIJYirWijYZzLhEiYmaZnW9/uzPszoI1lFYyep0bMtm8FvB77QvhRXvDM43YaGMER0TR1dngZIia9nGHxYw3LrNW7XP/WnMsritgaRQeV1QtLzwlPLOWV03eFu5SUpGY8KmwU5MLCt+aerTCLyYnK/xlshYK+MDaIWxP/uLoL1ZSWkZYXk5fJr2q/NzHfElLPDsflNgr3oNOAD9e7EwzgQ8Pg4zJ7MHFEAOyoka+u5w/R05yFZlVCmiskCRFHqeoq1I9LjEhelxGmoLZ/7991RPDQ5XqLV6ofzKMt35o2IJS0TA+Dw2jdAS2R7jIVvNzBzD6LnqxqvXtQ/sGnF1WtegOnG9C94Ma0SJlySZuTSTg9QRaw9B5DU1LlZ797HN8D6F1+aor2N0Dh5xvX/4Gl2ln/HOw8/IAAAAJcEhZcwAADsQAAA7EAZUrDhsAAABCSURBVEiJ7dYhEgAgEEJRVs3e/5x2B4vFvlucTyLxKmFbpSkFbI9bV8+fn1tSy999AwAAAAAAAAAA8AtQfn6j+r4fsXEUKxsy0BIAAAAASUVORK5CYII=";
        string red32 = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAFSmlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4KPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iWE1QIENvcmUgNS41LjAiPgogPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4KICA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIgogICAgeG1sbnM6ZGM9Imh0dHA6Ly9wdXJsLm9yZy9kYy9lbGVtZW50cy8xLjEvIgogICAgeG1sbnM6ZXhpZj0iaHR0cDovL25zLmFkb2JlLmNvbS9leGlmLzEuMC8iCiAgICB4bWxuczp0aWZmPSJodHRwOi8vbnMuYWRvYmUuY29tL3RpZmYvMS4wLyIKICAgIHhtbG5zOnBob3Rvc2hvcD0iaHR0cDovL25zLmFkb2JlLmNvbS9waG90b3Nob3AvMS4wLyIKICAgIHhtbG5zOnhtcD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyIKICAgIHhtbG5zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIgogICAgeG1sbnM6c3RFdnQ9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZUV2ZW50IyIKICAgZXhpZjpQaXhlbFhEaW1lbnNpb249IjMyIgogICBleGlmOlBpeGVsWURpbWVuc2lvbj0iMzIiCiAgIGV4aWY6Q29sb3JTcGFjZT0iMSIKICAgdGlmZjpJbWFnZVdpZHRoPSIzMiIKICAgdGlmZjpJbWFnZUxlbmd0aD0iMzIiCiAgIHRpZmY6UmVzb2x1dGlvblVuaXQ9IjIiCiAgIHRpZmY6WFJlc29sdXRpb249Ijk2LjAiCiAgIHRpZmY6WVJlc29sdXRpb249Ijk2LjAiCiAgIHBob3Rvc2hvcDpDb2xvck1vZGU9IjMiCiAgIHBob3Rvc2hvcDpJQ0NQcm9maWxlPSJzUkdCIElFQzYxOTY2LTIuMSIKICAgeG1wOk1vZGlmeURhdGU9IjIwMjItMDgtMDNUMTE6MDE6MjktMDQ6MDAiCiAgIHhtcDpNZXRhZGF0YURhdGU9IjIwMjItMDgtMDNUMTE6MDE6MjktMDQ6MDAiPgogICA8ZGM6dGl0bGU+CiAgICA8cmRmOkFsdD4KICAgICA8cmRmOmxpIHhtbDpsYW5nPSJ4LWRlZmF1bHQiPkljb25zPC9yZGY6bGk+CiAgICA8L3JkZjpBbHQ+CiAgIDwvZGM6dGl0bGU+CiAgIDx4bXBNTTpIaXN0b3J5PgogICAgPHJkZjpTZXE+CiAgICAgPHJkZjpsaQogICAgICBzdEV2dDphY3Rpb249InByb2R1Y2VkIgogICAgICBzdEV2dDpzb2Z0d2FyZUFnZW50PSJBZmZpbml0eSBQaG90byAxLjkuMiIKICAgICAgc3RFdnQ6d2hlbj0iMjAyMi0wOC0wM1QxMTowMToyOS0wNDowMCIvPgogICAgPC9yZGY6U2VxPgogICA8L3htcE1NOkhpc3Rvcnk+CiAgPC9yZGY6RGVzY3JpcHRpb24+CiA8L3JkZjpSREY+CjwveDp4bXBtZXRhPgo8P3hwYWNrZXQgZW5kPSJyIj8+hYzfmgAAAYFpQ0NQc1JHQiBJRUM2MTk2Ni0yLjEAACiRdZG7SwNBEIe/JIriG7SwsAiiqRKJClEbiwSNglokEYzaJJeXkMdxF5FgK9gGFEQbX4X+BdoK1oKgKIJYirWijYZzLhEiYmaZnW9/uzPszoI1lFYyep0bMtm8FvB77QvhRXvDM43YaGMER0TR1dngZIia9nGHxYw3LrNW7XP/WnMsritgaRQeV1QtLzwlPLOWV03eFu5SUpGY8KmwU5MLCt+aerTCLyYnK/xlshYK+MDaIWxP/uLoL1ZSWkZYXk5fJr2q/NzHfElLPDsflNgr3oNOAD9e7EwzgQ8Pg4zJ7MHFEAOyoka+u5w/R05yFZlVCmiskCRFHqeoq1I9LjEhelxGmoLZ/7991RPDQ5XqLV6ofzKMt35o2IJS0TA+Dw2jdAS2R7jIVvNzBzD6LnqxqvXtQ/sGnF1WtegOnG9C94Ma0SJlySZuTSTg9QRaw9B5DU1LlZ797HN8D6F1+aor2N0Dh5xvX/4Gl2ln/HOw8/IAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAA/SURBVEiJ7dahEQAwEAJByKT/lol58wVgModCsRYnUTVVIMmdVli3JOkUllcAAAAAAAAAAAB+Aeabuga4fd8fUhMNNmro7pAAAAAASUVORK5CYII=";
        string yellow32 = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAFQWlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4KPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iWE1QIENvcmUgNS41LjAiPgogPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4KICA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIgogICAgeG1sbnM6ZGM9Imh0dHA6Ly9wdXJsLm9yZy9kYy9lbGVtZW50cy8xLjEvIgogICAgeG1sbnM6ZXhpZj0iaHR0cDovL25zLmFkb2JlLmNvbS9leGlmLzEuMC8iCiAgICB4bWxuczp0aWZmPSJodHRwOi8vbnMuYWRvYmUuY29tL3RpZmYvMS4wLyIKICAgIHhtbG5zOnBob3Rvc2hvcD0iaHR0cDovL25zLmFkb2JlLmNvbS9waG90b3Nob3AvMS4wLyIKICAgIHhtbG5zOnhtcD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyIKICAgIHhtbG5zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIgogICAgeG1sbnM6c3RFdnQ9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9SZXNvdXJjZUV2ZW50IyIKICAgZXhpZjpQaXhlbFhEaW1lbnNpb249IjMyIgogICBleGlmOlBpeGVsWURpbWVuc2lvbj0iMzIiCiAgIGV4aWY6Q29sb3JTcGFjZT0iMSIKICAgdGlmZjpJbWFnZVdpZHRoPSIzMiIKICAgdGlmZjpJbWFnZUxlbmd0aD0iMzIiCiAgIHRpZmY6UmVzb2x1dGlvblVuaXQ9IjIiCiAgIHRpZmY6WFJlc29sdXRpb249Ijk2LjAiCiAgIHRpZmY6WVJlc29sdXRpb249Ijk2LjAiCiAgIHBob3Rvc2hvcDpDb2xvck1vZGU9IjMiCiAgIHBob3Rvc2hvcDpJQ0NQcm9maWxlPSJzUkdCIElFQzYxOTY2LTIuMSIKICAgeG1wOk1vZGlmeURhdGU9IjIwMjItMDgtMDNUMTE6MDEtMDQ6MDAiCiAgIHhtcDpNZXRhZGF0YURhdGU9IjIwMjItMDgtMDNUMTE6MDEtMDQ6MDAiPgogICA8ZGM6dGl0bGU+CiAgICA8cmRmOkFsdD4KICAgICA8cmRmOmxpIHhtbDpsYW5nPSJ4LWRlZmF1bHQiPkljb25zPC9yZGY6bGk+CiAgICA8L3JkZjpBbHQ+CiAgIDwvZGM6dGl0bGU+CiAgIDx4bXBNTTpIaXN0b3J5PgogICAgPHJkZjpTZXE+CiAgICAgPHJkZjpsaQogICAgICBzdEV2dDphY3Rpb249InByb2R1Y2VkIgogICAgICBzdEV2dDpzb2Z0d2FyZUFnZW50PSJBZmZpbml0eSBQaG90byAxLjkuMiIKICAgICAgc3RFdnQ6d2hlbj0iMjAyMi0wOC0wM1QxMTowMS0wNDowMCIvPgogICAgPC9yZGY6U2VxPgogICA8L3htcE1NOkhpc3Rvcnk+CiAgPC9yZGY6RGVzY3JpcHRpb24+CiA8L3JkZjpSREY+CjwveDp4bXBtZXRhPgo8P3hwYWNrZXQgZW5kPSJyIj8+QlHHJAAAAYFpQ0NQc1JHQiBJRUM2MTk2Ni0yLjEAACiRdZG7SwNBEIe/JIriG7SwsAiiqRKJClEbiwSNglokEYzaJJeXkMdxF5FgK9gGFEQbX4X+BdoK1oKgKIJYirWijYZzLhEiYmaZnW9/uzPszoI1lFYyep0bMtm8FvB77QvhRXvDM43YaGMER0TR1dngZIia9nGHxYw3LrNW7XP/WnMsritgaRQeV1QtLzwlPLOWV03eFu5SUpGY8KmwU5MLCt+aerTCLyYnK/xlshYK+MDaIWxP/uLoL1ZSWkZYXk5fJr2q/NzHfElLPDsflNgr3oNOAD9e7EwzgQ8Pg4zJ7MHFEAOyoka+u5w/R05yFZlVCmiskCRFHqeoq1I9LjEhelxGmoLZ/7991RPDQ5XqLV6ofzKMt35o2IJS0TA+Dw2jdAS2R7jIVvNzBzD6LnqxqvXtQ/sGnF1WtegOnG9C94Ma0SJlySZuTSTg9QRaw9B5DU1LlZ797HN8D6F1+aor2N0Dh5xvX/4Gl2ln/HOw8/IAAAAJcEhZcwAADsQAAA7EAZUrDhsAAABESURBVEiJ7dYhEgAgEELRXfX+N3YwqMW8FOeTSLxKSgprrICkcVsvX8+cEdHKd58AAAAAAAAAAAD8Apxvun+kI+m+7wtioxA0Q8wTIgAAAABJRU5ErkJggg==";

        public ButtonClass(string name, string text, string className, System.Drawing.Bitmap largeImage,
            System.Drawing.Bitmap smallImage, string toolTip)
        {
            Data = new PushButtonData(name, text, GetAssemblyName(), className);
            Data.ToolTip = toolTip;
            Data.LargeImage = BitmapToImageSource(largeImage);
            //Data.LargeImage = Base64ToImageSource(blue32);
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

        private Bitmap Base64ToBitmap(string base64String)
        {
            Bitmap bmpReturn = null;
            //Convert Base64 string to byte[]
            byte[] byteBuffer = Convert.FromBase64String(base64String);
            MemoryStream memoryStream = new MemoryStream(byteBuffer);

            memoryStream.Position = 0;

            bmpReturn = (Bitmap)Bitmap.FromStream(memoryStream);

            memoryStream.Close();
            memoryStream = null;
            byteBuffer = null;

            return bmpReturn;
        }


        private BitmapImage Base64ToImageSource(string base64String)
        {
            Bitmap bitmap = Base64ToBitmap(base64String);

            IntPtr hBitmap = bitmap.GetHbitmap();
            BitmapImage retval;

            try
            {
                retval = (BitmapImage)Imaging.CreateBitmapSourceFromHBitmap(
                             hBitmap,
                             IntPtr.Zero,
                             Int32Rect.Empty,
                             BitmapSizeOptions.FromEmptyOptions());
            }
            finally
            {}

            return retval;
        }
    }
}
