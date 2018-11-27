using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Autodesk.Revit.ApplicationServices;
using RVTApplication = Autodesk.Revit.ApplicationServices.Application;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace BARevitRoutines
{
    public partial class RequestHandler: IExternalEventHandler
    {
        UIControlledApplication uiApp = null;
        public RequestHandler(UIControlledApplication newUIApp)
        {
            uiApp = newUIApp;
        }
        public Request m_request = new Request();
        public Request Request
        {
            get { return m_request; }
        }
        public String GetName()
        {
            return "BA Revit Routines External Event";
        }
        public void Execute(UIApplication uiApp)
        {
            try
            {
                switch (Request.Take())
                {
                    case RequestId.None:
                        return;
                    case RequestId.CollectProjectLinksData:
                        BARevitRoutines.RoutineRequests.CollectProjectLinksData collectProjectLinksData = new RoutineRequests.CollectProjectLinksData(uiApp, "Collect Links Data");
                        break;
                }
            }
            finally
            {
                ;
            }
            return;
        }        
    }  
}
