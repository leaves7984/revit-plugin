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
    class EditElement : IExternalCommand
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
                        trans.Start();

                        // move element
                        XYZ moveVec = new XYZ(3, 3, 0);
                        ElementTransformUtils.MoveElement(doc, eleId, moveVec);

                        // rotate element
                        LocationPoint loc = ele.Location as LocationPoint;
                        XYZ p1 = loc.Point;
                        XYZ p2 = new XYZ(p1.X, p1.Y, p1.Z + 10);
                        Line axis = Line.CreateBound(p1, p2);
                        double angle = 30 * Math.PI / 180;
                        ElementTransformUtils.RotateElement(doc, eleId, axis, angle);
                        trans.Commit();

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
