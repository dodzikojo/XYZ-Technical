using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace xyzTechnicalRevit
{

    //Create an Addin/Plugin in Revit with the commands[separate commands for each task] to do the following tasks:
    //To list all the 3D views in the active project.
    //To create a text file containing the metadata of the elements that are present in each view, along with the element ID and layer name(describing which elements belongs to which layer)
    //To add a new parameter with a name “XYZ UniqueID” which holds the unique Id of the elements. [This should be visible in the properties panel for each element]
    //To export the selected element’s geometry to OBJ or SVF format.
    public class MainClass : IExternalApplication
    {

        public Result OnStartup(UIControlledApplication application)
        {
            //Add Tab to Revit UI
            string tabName = "XYZ Technical";

            //Create tab on revit UI
            application.CreateRibbonTab(tabName);


            // Add a new ribbon panel
            RibbonPanel ribbonPanel = application.CreateRibbonPanel(tabName, "XYZ Tools");

            //Add Ribbon to Revit UI
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            //Create buttons
            var getList3dViewsBtnData = new PushButtonData("getList3dViews", "List 3D Views", thisAssemblyPath, "xyzTechnicalRevit.commands.getList3dViewsCommand")
            {
                ToolTip = "List all the 3D views in the active project",
                LargeImage = new BitmapImage(new Uri("pack://application:,,,/xyzTechnicalRevit;component/Resources/icon-32.png")),
                Image = new BitmapImage(new Uri("pack://application:,,,/xyzTechnicalRevit;component/Resources/icon-16.png")),
            };

            var metadataExportBtnData = new PushButtonData("metadataExport", "Export Elements\nMetadata", thisAssemblyPath, "xyzTechnicalRevit.commands.metadataExportCommand")
            {
                ToolTip = "Creates a text file containing the metadata of the elements that are present in each view, along with the element ID and layer name(describing which elements belongs to which layer)",
                LargeImage = new BitmapImage(new Uri("pack://application:,,,/xyzTechnicalRevit;component/Resources/icon-32.png")),
                Image = new BitmapImage(new Uri("pack://application:,,,/xyzTechnicalRevit;component/Resources/icon-16.png")),
            };

            var addParameterBtnData = new PushButtonData("addParameter", "Create XYZ UniqueID\nParameter", thisAssemblyPath, "xyzTechnicalRevit.commands.addParameterCommand")
            {
                ToolTip = "Adds a new parameter with a name “XYZ UniqueID” which holds the unique Id of the elements",
                LargeImage = new BitmapImage(new Uri("pack://application:,,,/xyzTechnicalRevit;component/Resources/icon-32.png")),
                Image = new BitmapImage(new Uri("pack://application:,,,/xyzTechnicalRevit;component/Resources/icon-16.png")),
            };

            var exportGeometryBtnData = new PushButtonData("exportGeometry", "Export Geometry", thisAssemblyPath, "xyzTechnicalRevit.commands.exportGeometryCommand")
            {
                ToolTip = "Export the selected element’s geometry",
                LargeImage = new BitmapImage(new Uri("pack://application:,,,/xyzTechnicalRevit;component/Resources/icon-32.png")),
                Image = new BitmapImage(new Uri("pack://application:,,,/xyzTechnicalRevit;component/Resources/icon-16.png")),
            };


            PushButton getList3dViewsBtn = ribbonPanel.AddItem(getList3dViewsBtnData) as PushButton;
            ribbonPanel.AddSeparator();
            PushButton metadataExportBtn = ribbonPanel.AddItem(metadataExportBtnData) as PushButton;
            ribbonPanel.AddSeparator();
            PushButton addParameterBtn = ribbonPanel.AddItem(addParameterBtnData) as PushButton;
            ribbonPanel.AddSeparator();
            PushButton exportGeometryBtn = ribbonPanel.AddItem(exportGeometryBtnData) as PushButton;


            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

       
    }
}
