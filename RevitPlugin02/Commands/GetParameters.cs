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
    class GetParameters : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try {
                // pick object
                Reference pickedObj = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
                if (pickedObj != null) {
                    // Retrieve Element
                    ElementId eleId = pickedObj.ElementId;
                    Element ele = doc.GetElement(eleId);

                    //get paramenter
                    Parameter param = ele.LookupParameter("Head Height");
                    InternalDefinition paramDef = param.Definition as InternalDefinition;

                    TaskDialog.Show("Parameters ", string.Format("{0} parameter of type {1} with builtinparameter {2}",
                        paramDef.Name,
                        paramDef.UnitType,
                        paramDef.BuiltInParameter));

                    
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
