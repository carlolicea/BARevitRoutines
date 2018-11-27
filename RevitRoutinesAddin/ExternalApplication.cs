#region Namespaces
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Autodesk.Revit.ApplicationServices;
using RVTApplication = Autodesk.Revit.ApplicationServices.Application;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
#endregion

namespace RevitRoutinesAddin
{
    public class Application : IExternalApplication
    {
        internal static Application thisApp = null;
        public static UIControlledApplication _cachedUiCtrApp;
        public RoutinesToRun stuffToRun = null;

        public Result OnStartup(UIControlledApplication application)
        {
            _cachedUiCtrApp = application;
            thisApp = this;
            application.ControlledApplication.ApplicationInitialized += ApplicationInitialized;
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        public void ApplicationInitialized(object sender, ApplicationInitializedEventArgs args)
        {
            RequestHandler handler = new RequestHandler(_cachedUiCtrApp);
            ExternalEvent exEvent = ExternalEvent.Create(handler);

            UIApplication uiApp = sender as UIApplication;
            stuffToRun = new RoutinesToRun(uiApp, exEvent, handler);
        }
    }
}