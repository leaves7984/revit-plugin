using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;

namespace RevitPlugin02
{
    [TransactionAttribute(TransactionMode.Manual)]
    class TagView : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // get UIdocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //get document
            Document doc = uidoc.Document;

            //Tag parameters
            TagMode tmode = TagMode.TM_ADDBY_CATEGORY;
            TagOrientation torient = TagOrientation.Horizontal;

            List<BuiltInCategory> cats = new List<BuiltInCategory>();
            cats.Add(BuiltInCategory.OST_Windows);
            cats.Add(BuiltInCategory.OST_Doors);

            ElementMulticategoryFilter filter = new ElementMulticategoryFilter(cats);
            IList<Element> telements = new FilteredElementCollector(doc, doc.ActiveView.Id)
                .WherePasses(filter)
                .WhereElementIsNotElementType()
                .ToElements();
            
            try
            {
                using (Transaction trans = new Transaction(doc, "create plan view"))
                {
                    trans.Start();
                    // tag elements
                    foreach (Element ele in telements) {
                        Reference refe = new Reference(ele);
                        LocationPoint loc = ele.Location as LocationPoint;
                        XYZ point = loc.Point;
                        IndependentTag tag = IndependentTag.Create(doc, doc.ActiveView.Id, refe, true, tmode, torient, point);
                    }
                    trans.Commit();
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

