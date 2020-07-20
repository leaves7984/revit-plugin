using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.ApplicationServices;
using System.IO;
using System.Net;
using System.Configuration;

namespace RevitPlugin02
{
    [TransactionAttribute(TransactionMode.Manual)]
    class exportNWC : IExternalCommand
    {
        Application m_rvtApp;
        Document m_rvtDoc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication rvtUIAPP = commandData.Application;
            UIDocument uidoc = rvtUIAPP.ActiveUIDocument;
            m_rvtDoc = uidoc.Document;
            m_rvtApp = rvtUIAPP.Application;
            Configuration config = ConfigurationManager.OpenExeConfiguration
                   (System.Reflection.Assembly.GetExecutingAssembly().Location);
            string familyRepo = config.AppSettings.Settings["Family_repo_folder"].Value;
            string serverAPI = config.AppSettings.Settings["Server_API"].Value;
            try
            {
                string[] filePaths = Directory.GetFiles(@familyRepo, "*.rfa");
                foreach (string path in filePaths)
                {
                    Document doc = m_rvtApp.OpenDocumentFile(path);
                    string filename = Path.GetFileName(path).Split('.')[0] + ".nwc";
                    try
                    {
                        NavisworksExportOptions options = new NavisworksExportOptions();
                        options.ExportScope = NavisworksExportScope.Model;
                        options.ViewId = uidoc.ActiveView.Id;

                        //doc.ActiveView.DetailLevel = ViewDetailLevel.Fine;
                        //doc.ActiveView.DisplayStyle = DisplayStyle.Realistic;
                        doc.Export(@familyRepo, filename, options);
         
                    }
                    catch (Exception e)
                    {
                        TaskDialog.Show("Error", filename + ":\n" + e.Message);
                    }

                }

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                message = e.Message;
                return Result.Failed;
            }
        }
     
    }
}
