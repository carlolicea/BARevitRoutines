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
using Autodesk.Revit.UI;
#endregion

namespace BARevitRoutines
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,ref string message,ElementSet elements)
        {
            MessageBox.Show("Executed External Command");
            try
            {                              
                return Result.Succeeded;
            }
            catch(Exception ExecuteException)
            {
                MessageBox.Show(ExecuteException.ToString());
                return Result.Failed;
            }
            
        }
    }  
}