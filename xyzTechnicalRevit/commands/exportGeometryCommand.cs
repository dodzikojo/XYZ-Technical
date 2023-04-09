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

                UIDocument uidoc = commandData.Application.ActiveUIDocument;

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

                // Create the file name
                //string fileName = "selected_elements_5.obj";

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

        //public void ExportToObj(Element element, string filePath)
        //{
        //    Document doc = element.Document;

        //    OBJExportOptions oBJExportOptions = new OBJExportOptions();

        //    GeometryElement geometryElement = element.get_Geometry(new Options());
        //    if (geometryElement == null) return;

        //    // Create a new OBJ file
        //    using (StreamWriter writer = new StreamWriter(filePath))
        //    {
        //        foreach (GeometryObject geoObj in geometryElement)
        //        {
        //            // Export only solid geometry
        //            if (geoObj is Solid solid && solid.Volume > 0)
        //            {
        //                // Export each face of the solid geometry
        //                foreach (Face face in solid.Faces)
        //                {
        //                    Mesh mesh = face.Triangulate();
        //                    if (mesh == null) continue;

        //                    // Export the vertices of the mesh
        //                    foreach (XYZ vertex in mesh.Vertices)
        //                    {
        //                        writer.WriteLine("v {0} {1} {2}", vertex.X, vertex.Y, vertex.Z);
        //                    }

        //                    // Export the indices of the mesh triangles
        //                    for (int i = 0; i < mesh.NumTriangles; i++)
        //                    {
        //                        MeshTriangle triangle = mesh.get_Triangle(i);
        //                        writer.WriteLine("f {0} {1} {2}", triangle.VertexIndex1 + 1, triangle.VertexIndex2 + 1, triangle.VertexIndex3 + 1);
        //                    }
        //                }
        //            }
        //        }
        //    }
        //}

    }
}
