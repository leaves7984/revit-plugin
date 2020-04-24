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
    class ViewFilter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // get UIdocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            //get document
            Document doc = uidoc.Document;

            //create filter
            List<ElementId> cats = new List<ElementId>();
            cats.Add(new ElementId(BuiltInCategory.OST_Sections));

            ElementParameterFilter filter = new ElementParameterFilter(ParameterFilterRuleFactory.CreateContainsRule(new ElementId(BuiltInParameter.VIEW_NAME),"wip", false));

            try
            {
                using (Transaction trans = new Transaction(doc, "create plan view"))
                {
                    trans.Start();
                    //apply filter
                    ParameterFilterElement filterElement = ParameterFilterElement.Create(doc, "My First Filter", cats, filter);
                    doc.ActiveView.AddFilter(filterElement.Id);
                    doc.ActiveView.SetFilterVisibility(filterElement.Id, false);
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

