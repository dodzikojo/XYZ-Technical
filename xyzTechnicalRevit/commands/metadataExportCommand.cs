using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace xyzTechnicalRevit.commands
{
    //To create a text file containing the metadata of the elements that are present in each view, along with the element ID and layer name(describing which elements belongs to which layer)
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.ReadOnly)]
    public class metadataExportCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get UIDocument
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            //Get Document
            Document doc = uidoc.Document;

            WriteMetadataToFile(doc);

            return Result.Succeeded;
        }

        public void WriteMetadataToFile(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(View3D));
            List<View3D> views = collector.Cast<View3D>().ToList();
            

            foreach (View view in views)
            {
                if (!view.IsTemplate)
                {
                    string fileName = view.Name + ".txt";
                    string filePath = @"D:\Downloads\metdata\" + fileName;

                    using (StreamWriter sw = File.CreateText(filePath))
                    {
                        sw.WriteLine("Metadata for View: " + view.Name);
                        sw.WriteLine("");

                        FilteredElementCollector viewCollector = new FilteredElementCollector(doc, view.Id);
                        ICollection<ElementId> elementIds = viewCollector.ToElementIds();


                        foreach (ElementId elementId in elementIds)
                        {
                            Element element = doc.GetElement(elementId);

                            if (element.Category != null)
                            {
                                string metadata = "Element ID: " + elementId.IntegerValue.ToString() + ", Layer Name: " + element.Category.Name.ToString();
                                sw.WriteLine(metadata);
                            }
                            else
                            {
                                
                            }
                            
                        }
                    }
                }
                
            }
        }
    }
}
