using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using wf = System.Windows.Forms;

// Navisworks API References
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using System.Windows.Forms;
using Application = Autodesk.Navisworks.Api.Application;
using System.Diagnostics;
using System.IO;

//To display the origin coordinates of the model in a pop-up message box/dialog.
//To create a text file containing the element ID of all the elements in the model
//To list only the element ID’s of the objects which are present inside the section box[When a section box is created]. 
//To export the selected element’s geometry to OBJ or FBX format.

namespace xyzTechnicalNavisworks
{


    // plugin constructor - custom tab attributes
    [Plugin("xyzTechnicalNavisworks", "xyzTechnicalNavisworks", DisplayName = "xyzTechnicalNavisworks")]
    // xaml file - layout of custom ribbon (panel & buttons)
    [RibbonLayout("xyzTechnicalNavisworks.xaml")]
    // ribbon tab ID from xaml file
    [RibbonTab("xyzTechnicalNavisworksTab")]
    // ribbon button ID & icon files
    [Command("Button_One", Icon = "icon-16.png", LargeIcon = "icon-32.png")]
    [Command("Button_Two", Icon = "icon-16.png", LargeIcon = "icon-32.png")]
    // split button with (btn-3 & btn-4) 
    //[Command("Split_Button_1")]
    [Command("Button_Three", Icon = "icon-16.png", LargeIcon = "icon-32.png")]
    [Command("Button_Four", Icon = "icon-16.png", LargeIcon = "icon-32.png")]
    public class xyzTechnicalNavisworksCommandHandler : CommandHandlerPlugin
    {
        public static string newLineValue { get; set; } = string.Empty;
        public static string previousLineValue { get; set; } = string.Empty;
        public override int ExecuteCommand(string name, params string[] parameters)
        {
            switch (name)
            {
                case "Button_One":
                    GetBoundingBoxCoordinates();
                    break;
                case "Button_Two":
                    CreateTextFile();
                    break;
                case "Button_Three":
                    
                    break;
                case "Button_Four":
                    
                    break;
            }
            return 0;
        }




        public void GetBoundingBoxCoordinates()
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            ModelItemCollection selection = doc.CurrentSelection.SelectedItems;
            BoundingBox3D bbox = selection.BoundingBox();
            double xMin = bbox.Min.X;
            double yMin = bbox.Min.Y;
            double zMin = bbox.Min.Z;
            double xMax = bbox.Max.X;
            double yMax = bbox.Max.Y;
            double zMax = bbox.Max.Z;
            string message = "The bounding box coordinates are: (" + xMin + ", " + yMin + ", " + zMin + ") to (" + xMax + ", " + yMax + ", " + zMax + ")";
            System.Windows.Forms.MessageBox.Show(message);
        }


        public void CreateTextFile()
        {
            Document doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
            var items = doc.Models.RootItems.ToList();


            SaveFileDialog saveFileDialog = new SaveFileDialog();
            string fullPath = string.Empty;
            string fileName = string.Empty;

            saveFileDialog.Filter = "Text files (*.txt)|*.txt";


            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string selectedPath = Path.GetDirectoryName(saveFileDialog.FileName);
                fileName = Path.GetFileName(saveFileDialog.FileName);
                fullPath = Path.Combine(selectedPath, fileName);

                using (StreamWriter sw = File.CreateText(fullPath))
                {
                    foreach (ModelItem item in items)
                    {
                        try
                        {
                            if (item.Children != null)
                            {
                                LoopThroughNestedLists(item.Children, sw);
                            }

                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.ToString()); 
                        }

                    }
                }

            }

            
        }

        public static void LoopThroughNestedLists(ModelItemEnumerableCollection modelItemChildren, StreamWriter sw)
        {
            foreach (ModelItem child in modelItemChildren)
            {
                if (child.Children != null)
                {
                    var propertyCategory = child.PropertyCategories.FindCategoryByDisplayName("Element ID");

                    if (propertyCategory != null)
                    {
                        if (propertyCategory.DisplayName == "Element ID")
                        {
                            foreach (DataProperty property in propertyCategory.Properties)
                            {
                                if (property.DisplayName == "Value")
                                {
                                    newLineValue = child.DisplayName + ": " + property.Value.ToDisplayString();
                                    if (newLineValue != previousLineValue)
                                    {
                                        sw.WriteLine(newLineValue);
                                    }
                                    previousLineValue = newLineValue;
                                }
                            }
                        }
                    }
                    // Call the recursive method with the current item
                    LoopThroughNestedLists(child.Children, sw);


                }
            }

        }











    }


}
