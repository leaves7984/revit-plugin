using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace RevitPlugin02
{
    class ExternalApplication : IExternalApplication
    {
        Result IExternalApplication.OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        Result IExternalApplication.OnStartup(UIControlledApplication application)
        {
            // create ribbon tab
    
            application.CreateRibbonTab("Cost");
            string path = Assembly.GetExecutingAssembly().Location;
            PushButtonData button = new PushButtonData("button1", "Cost", path, "RevitPlugin02.GetAllFamily");
            RibbonPanel panel = application.CreateRibbonPanel("Cost", "Cost");
            panel.AddItem(button);
            
            return Result.Succeeded;
        }
    }
}
