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
    class PlaceFamily : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // get UIdocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //get document
            Document doc = uidoc.Document;

            //get family symbol
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            FamilySymbol symbol = collector.OfClass(typeof(FamilySymbol))
                .WhereElementIsElementType()
                .Cast<FamilySymbol>()
                .First(x => x.Name == "1525 x 762mm");
            /*
             * FamilySymbol symbol = null;
            foreach (Element ele in symbols) {
                if (ele.Name == "1525 x 762mm") {
                    symbol = ele as FamilySymbol;
                    break;
                }
            }*/
            try
            {
                using (Transaction trans = new Transaction(doc, "Place Family")) {
                    trans.Start();
                    if (!symbol.IsActive) {
                        symbol.Activate();
                    }
                    doc.Create.NewFamilyInstance(new XYZ(0, 0, 0), symbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                    trans.Commit();
                }

                return Result.Succeeded;

            }
            catch (Exception e) {
                message = e.Message;
                return Result.Failed;
            }
        }
    }
}
