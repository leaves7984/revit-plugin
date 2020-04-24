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
    class ProjectRay : IExternalCommand
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
                    //Retrieve Element
                    ElementId eleId = pickedObj.ElementId;
                    Element ele = doc.GetElement(eleId);

                    //Project Ray
                    LocationPoint locp = ele.Location as LocationPoint;
                    XYZ p1 = locp.Point;

                    // Ray
                    XYZ rayd = new XYZ(0, 0, 1);
                    ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_Roofs);
                    ReferenceIntersector refi = new ReferenceIntersector(filter, FindReferenceTarget.Face, (View3D)doc.ActiveView);
                    ReferenceWithContext refc = refi.FindNearest(p1, rayd);
                    Reference reference = refc.GetReference();
                    XYZ intpoing = reference.GlobalPoint;
                    Double dist = p1.DistanceTo(intpoing);

                    TaskDialog.Show("Ray", string.Format("Distance to roof {0}", dist));
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
