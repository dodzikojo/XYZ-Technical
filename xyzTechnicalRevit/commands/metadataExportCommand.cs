using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Xml.Linq;
using View = Autodesk.Revit.DB.View;

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

            FolderBrowserDialog folderBrowserDialog1 = new FolderBrowserDialog();
            folderBrowserDialog1.Description = "Select or create new folder to save the file(s)";

            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string folderPath = string.Empty;
                folderPath = folderBrowserDialog1.SelectedPath;

                try
                {
                    foreach (View view in views)
                    {
                        if (!view.IsTemplate)
                        {
                            string fileName = view.Name + ".txt";

                            using (StreamWriter sw = File.CreateText(Path.Combine(folderPath, fileName)))
                            {
                                sw.WriteLine("Metadata for View: " + view.Name);
                                sw.WriteLine("");

                                FilteredElementCollector viewCollector = new FilteredElementCollector(doc, view.Id);
                                ICollection<ElementId> elementIds = viewCollector.ToElementIds();

                                // Loop through all elements and extract their parameters
                                foreach (ElementId elementId in elementIds)
                                {
                                    Element element = doc.GetElement(elementId);

                                    // Get the element's parameters
                                    ParameterSet parameters = element.Parameters;

                                    // Loop through all parameters and extract their values
                                    foreach (Parameter parameter in parameters)
                                    {
                                        string paramName = parameter.Definition.Name;
                                        string paramValue = parameter.AsValueString();

                                        // Do something with the parameter name and value
                                        //Debug.WriteLine("Element ID: {0}, Parameter Name: {1}, Parameter Value: {2}", element.Id.ToString(), paramName, paramValue);
                                        sw.WriteLine("Element ID: {0}, Parameter Name: {1}, Parameter Value: {2}", element.Id.ToString(), paramName, paramValue);
                                    }
                                }

                            }


                        }


                    }
                }
                catch (UnauthorizedAccessException unAuthExp)
                {
                    TaskDialog.Show("Info", unAuthExp.Message);
                }
                catch (Exception ex) 
                { 
                    Debug.WriteLine(ex.ToString()); 
                }


            }
        }
    }
}
