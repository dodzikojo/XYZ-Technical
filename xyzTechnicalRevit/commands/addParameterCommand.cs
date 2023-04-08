using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;

namespace xyzTechnicalRevit.commands
{
    //To add a new parameter with a name “XYZ UniqueID” which holds the unique Id of the elements. [This should be visible in the properties panel for each element]

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class addParameterCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication application = commandData.Application;
            Document document = application.ActiveUIDocument.Document;

            using (Transaction t = new Transaction(document, "Create New Parameter"))
            {
                t.Start();
                CreateSharedParameter(document);
                AssignUniqueIds(document);
                t.Commit();
            }

            return Result.Succeeded;
        }




        public static void CreateSharedParameter(Document doc)
        {
            // Get a reference to the shared parameter file
            DefinitionFile sharedParameterFile = doc.Application.OpenSharedParameterFile();

            if (sharedParameterFile == null)
            {
                // create shared parameter file
                String modulePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                String paramFile = modulePath + "\\SharedParameters.txt";
                if (File.Exists(paramFile))
                {
                    File.Delete(paramFile);
                }
                FileStream fs = File.Create(paramFile);
                fs.Close();

                // Assign Shared Parameter File
                doc.Application.SharedParametersFilename = paramFile;

            }
            else
            {

                sharedParameterFile = doc.Application.OpenSharedParameterFile();

                //Check if the parameter already exist in the file.
                DefinitionGroup sharedParameterGroup = sharedParameterFile.Groups.get_Item("XYZ");
                Definition sharedParameter = null;
                if (sharedParameterGroup == null)
                {
                    // Create a new shared parameter group
                    sharedParameterGroup = sharedParameterFile.Groups.Create("XYZ");

                }
                else
                {
                    //CheckBox if the parameter name exist
                    sharedParameter = sharedParameterGroup.Definitions.get_Item("XYZ UniqueID");

                    if (sharedParameter == null)
                    {
                        // If not existing, create the shared parameter
                        ExternalDefinitionCreationOptions externalDefinitionCreationOptions = new ExternalDefinitionCreationOptions("XYZ UniqueID", SpecTypeId.String.Text);
                        externalDefinitionCreationOptions.UserModifiable = true;
                        externalDefinitionCreationOptions.Visible = true;
                        externalDefinitionCreationOptions.Description = "Unique ID of the element";

                        sharedParameter = sharedParameterGroup.Definitions.Create(externalDefinitionCreationOptions);
                    }
                }



                Categories categories = doc.Settings.Categories;

                List<string> list = getListOfCategories();
                CategorySet categorySet = new CategorySet();

                foreach (Category category in categories)
                {
                    if (!list.Contains(category.Name))
                    {
                        categorySet.Insert(category);
                    }
                }


                try
                {
                    assignCategorySet(doc, sharedParameter, categorySet);
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException inValidOpEx)
                {
                    Debug.WriteLine(inValidOpEx.Message);

                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);

                }
            }

        }

        public static void assignCategorySet(Document doc, Definition sharedParameter, CategorySet categorySet)
        {

            InstanceBinding binding = doc.Application.Create.NewInstanceBinding(categorySet);
            doc.ParameterBindings.Insert(sharedParameter, binding, BuiltInParameterGroup.PG_TEXT);


        }

        public static List<string> getListOfCategories()
        {
            string filePath = @"D:\Downloads\metdata\unassign.txt";
            List<string> lines = new List<string>();

            try
            {
                using (StreamReader sr = new StreamReader(filePath))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        lines.Add(line);
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("The file could not be read:");
                Console.WriteLine(e.Message);
            }

            return lines;
        }

        public void AssignUniqueIds(Document doc)
        {
            // Get all elements in the document
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> allElements = collector
                .WhereElementIsNotElementType()
                .ToElements();

            // Loop through each element and assign its ElementId to a parameter called "XYZ UniqueID"
            foreach (Element element in allElements)
            {
                Parameter parameter = element.LookupParameter("XYZ UniqueID");
                if (parameter != null)
                {
                    parameter.Set(element.Id.ToString());
                }
            }
        }



    }



}
