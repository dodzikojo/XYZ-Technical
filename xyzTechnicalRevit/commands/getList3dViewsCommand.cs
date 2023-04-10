using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace xyzTechnicalRevit.commands
{
    //To list all the 3D views in the active project.
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.ReadOnly)]
    public class getList3dViewsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            //Get Document
            Document doc = uidoc.Document;


            ListAll3DViews(doc);


            return Result.Succeeded;
        }

        /// <summary>
        /// Get all 3D views in the active project.
        /// </summary>
        /// <param name="doc"></param>
        public void ListAll3DViews(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(View3D));
            List<View3D> views = collector.Cast<View3D>().ToList();

            Window window = new Window();
            window.Title = "List of 3D Views";
            window.Width = 250;
            window.Height = 300;

            ListBox listBox = new ListBox();
            listBox.ItemsSource = views;
            listBox.DisplayMemberPath = "Name";

            window.Content = listBox;
            window.ShowDialog();

        }
    }
}
