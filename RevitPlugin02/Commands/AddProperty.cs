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
            m_rvtDoc = uidoc.Document;
            m_rvtApp= rvtUIAPP.Application;
            try
            {
                string[] filePaths = Directory.GetFiles(@"C:\Users\wanghp18\Desktop\family-repo", "*.rfa");
                foreach (string path in filePaths) {
                    Document doc = m_rvtApp.OpenDocumentFile(path);
                    if (doc.IsFamilyDocument) {
                        Family f = doc.OwnerFamily;
                        FamilyManager manager = doc.FamilyManager;
                        string types = "";
                        FamilyTypeSet familyTypes = manager.Types;
                        FamilyTypeSetIterator familyTypeSetIterator = familyTypes.ForwardIterator();
                        familyTypeSetIterator.Reset();
                        FamilyParameter keynote = manager.get_Parameter("Keynote");
                        using (Transaction trans = new Transaction(doc, "SET_PARAM")) {
                            trans.Start();
                            String note = HttpGET("http://192.168.1.159:8787/GetFamilyKeynote?path=" + path);
                            manager.Set(keynote, note);
                            trans.Commit();
                        }
                        doc.Save();

                        while (familyTypeSetIterator.MoveNext()) {
                            FamilyType type = familyTypeSetIterator.Current as FamilyType;
                            types += "\n" + type.Name;
                            string value = type.AsString(manager.get_Parameter("Keynote"));
                            /*
                             * TaskDialog.Show("value", value);
                             */

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
