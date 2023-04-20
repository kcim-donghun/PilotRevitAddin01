#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Reflection;
using System.Windows;
using System.Windows.Media.Imaging;

#endregion

namespace PilotRevitAddin01
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            // 탭 추가 
            a.CreateRibbonTab("KCIM");

            // 리본 추가 
            RibbonPanel panel = a.CreateRibbonPanel("KCIM", "Pilot01");

            // DLL 위치 
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            // 버튼 추가 
            if (panel.AddItem(new PushButtonData("WinForm01", "WinForm01", thisAssemblyPath, "PilotRevitAddin01.Command"))
                is Autodesk.Revit.UI.PushButton button)
            {
                button.ToolTip = "WinForm01";

                // 버튼 이미지추가 
                button.LargeImage = convertFromBitmap(Properties.Resources.Image1);
                button.Image = convertFromBitmap(Properties.Resources.Image1);
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }

        // 프로젝트 속성 리소스 이미지 불러오는 함수
        BitmapSource convertFromBitmap(Bitmap bitmap)
        {
            return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                bitmap.GetHbitmap(),
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());
        }
    }
}
