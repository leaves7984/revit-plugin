using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitPlugin02
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class Class1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            // Get Document
            Document doc = uidoc.Document;
          
            try {
                // Pick Object
                Reference pickedObj = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);

                //Retrieve Element
                ElementId eleId = pickedObj.ElementId;

                Element ele = doc.GetElement(eleId);

                //Get Element Type
                ElementId etypeId = ele.GetTypeId();
                ElementType eType = doc.GetElement(etypeId) as ElementType;


                // Display Element Id
                if (pickedObj != null)
                {
                    // Dialog Box
                    TaskDialog.Show("El;ement Classification", eleId.ToString() + Environment.NewLine
                        + "Category: " + ele.Category.Name + Environment.NewLine
                        + "Instance: " + ele.Name + Environment.NewLine
                        + "Symbol: " + eType.Name + Environment.NewLine
                        + "Family: " + eType.FamilyName);
                }
            }catch(Exception e)
            {
                message = e.Message;
                return Result.Failed;
            }
            
            return Result.Succeeded;
        }
    }
}
