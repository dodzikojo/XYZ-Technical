using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using SaveFileDialog = System.Windows.Forms.SaveFileDialog;
using View = Autodesk.Revit.DB.View;

namespace xyzTechnicalRevit.commands
{
    //To export the selected element’s geometry to OBJ or SVF format.
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class exportGeometryCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the current document
            Document doc = commandData.Application.ActiveUIDocument.Document;
            View activeView = commandData.Application.ActiveUIDocument.ActiveView;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            string fullPath = string.Empty;
            string fileName = string.Empty;

            saveFileDialog.Filter = "Obj files (*.obj)|*.obj";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string selectedPath = Path.GetDirectoryName(saveFileDialog.FileName);
                fileName = Path.GetFileName(saveFileDialog.FileName);
                fullPath = selectedPath;

                // Do something with the full path
                View3D active3dView = null;

                try
                {
                    active3dView = (View3D)activeView;
                }
                catch (Exception)
                {
                    TaskDialog.Show("Info", "Try again while in a 3D view");
                }

                FilteredElementCollector viewCollector = new FilteredElementCollector(doc, active3dView.Id);
                var allElement = viewCollector.WhereElementIsNotElementType().ToElements();

                // Get the selected elements
                IList<ElementId> selectedElementIds = (IList<ElementId>)commandData.Application.ActiveUIDocument.Selection.GetElementIds();

                var collector = new FilteredElementCollector(doc);
                var viewFamilyType = collector.OfClass(typeof(ViewFamilyType)).Cast<ViewFamilyType>()
                  .FirstOrDefault(x => x.ViewFamily == ViewFamily.ThreeDimensional);

                View3D view3D = null;

                using (Transaction t = new Transaction(doc, "Create Temporary View and Export OBJ"))
                {
                    t.Start();
                    view3D = View3D.CreateIsometric(doc, viewFamilyType.Id);
                    try
                    {
                        IList<ElementId> elementToHide = new List<ElementId>();
                        foreach (Element element in allElement)
                        {
                            if (!selectedElementIds.Contains(element.Id) && element.CanBeHidden(view3D))
                            {
                                elementToHide.Add(element.Id);
                            }
                        }
                        view3D.HideElements(elementToHide);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                    }

                    t.Commit();
                }


                commandData.Application.ActiveUIDocument.ActiveView = view3D;
                ElementId viewID = view3D.Id;

                //Create the export options
                OBJExportOptions options = new OBJExportOptions();

                options.ViewId = viewID;

                // Export the selected elements to the OBJ file
                doc.Export(fullPath, fileName, options);

                commandData.Application.ActiveUIDocument.ActiveView = active3dView;


                using (Transaction t = new Transaction(doc, "Delete Temporary View"))
                {
                    t.Start();
                    View e = doc.GetElement(viewID) as View;
                    doc.Delete(e.Id);
                    t.Commit();
                }

                //Display a message box to confirm the export
                TaskDialog.Show("Export Elements", "Selected elements have been exported to " + fileName);
            }
            else
            {
                TaskDialog.Show("Info", "Export Cancelled");
            }

            return Result.Succeeded;
        }


    }
}
