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
    class ChangeLocation : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // pick object
                Reference pickedObj = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
                if (pickedObj != null)
                {
                    // Retrieve Element
                    ElementId eleId = pickedObj.ElementId;
                    Element ele = doc.GetElement(eleId);

                    

                    using (Transaction trans = new Transaction(doc, "Place Family"))
                    {
                        
                        //set location
                        LocationPoint locp = ele.Location as LocationPoint;

                        if (locp != null) {
                            trans.Start();
                            XYZ loc = locp.Point;
                            XYZ newloc = new XYZ(loc.X + 3, loc.Y, loc.Z);

                            locp.Point = newloc;
                            trans.Commit();
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
    }
}
