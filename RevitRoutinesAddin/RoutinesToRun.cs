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
using Autodesk.Revit.UI;

namespace BARevitRoutines
{
    public class RoutinesToRun
    {
        private RequestHandler m_Handler;
        private ExternalEvent m_ExEvent;
        private UIApplication uiApp;

        public RoutinesToRun(UIApplication exUiApp, ExternalEvent exEvent, RequestHandler handler)
        {
            m_Handler = handler;
            m_ExEvent = exEvent;
            uiApp = exUiApp;

            m_ExEvent.Raise();
            MakeRequest(RequestId.CollectProjectLinksData);
        }

        private void MakeRequest(RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();

        }
    }
}
