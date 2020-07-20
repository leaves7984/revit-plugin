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
    class AddProperty : IExternalCommand
    {
        Application m_rvtApp;
        Document m_rvtDoc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication rvtUIAPP = commandData.Application;
            UIDocument uidoc = rvtUIAPP.ActiveUIDocument;
            Configuration config = ConfigurationManager.OpenExeConfiguration
                    (System.Reflection.Assembly.GetExecutingAssembly().Location);
            string familyRepo = config.AppSettings.Settings["Family_repo_folder"].Value;
            string serverAPI = config.AppSettings.Settings["Server_API"].Value;
            m_rvtDoc = uidoc.Document;
            m_rvtApp= rvtUIAPP.Application;
            try
            {
                
                string[] filePaths = Directory.GetFiles(@familyRepo, "*.rfa");
                foreach (string path in filePaths) {
                    Document doc = m_rvtApp.OpenDocumentFile(path);
                    if (doc.IsFamilyDocument) {

                        try {
                            Family f = doc.OwnerFamily;
                            FamilyManager manager = doc.FamilyManager;
                            FamilyParameter keynote = null;
                            
                            String note = HttpGET(serverAPI + "GetFamilyKeynote?path=" + path);
                            keynote = manager.get_Parameter(BuiltInParameter.KEYNOTE_PARAM);
                            if (keynote != null)
                            {
                                
                                using (Transaction trans = new Transaction(doc, "SET_PARAM"))
                                {
                                    trans.Start();
                                    if (manager.Types.Size == 0)
                                    {
                                        manager.NewType("Type 1");
                                    }
                                    manager.SetFormula(keynote, null);
                                    manager.Set(keynote, note);
                                    trans.Commit();
                                }
                                doc.Save();
                            }
                        }
                        catch (Exception e)
                        {
                            TaskDialog.Show("Error2", e.Message);
                        }



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
        public static string HttpGET(string url)
        {
            var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
            httpWebRequest.Method = "GET"; 
            var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
            {
                var result = streamReader.ReadToEnd();
                return result;
            }


        }
    }
}
