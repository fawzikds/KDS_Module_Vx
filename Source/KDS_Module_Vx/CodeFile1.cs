using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;



public class SampleCreateSharedParameter : IExternalCommand
{
    #region // Execute Region of Code (Invoked by KDS Ribbon) //
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {
        return Result.Succeeded;
    }
    #endregion  //End Of Execute Region of Code (Invoked by KDS Ribbon) //
    public void CreateSampleSharedParameters(Document doc, Autodesk.Revit.ApplicationServices.Application app, UIApplication uiapp)
    {
        string mySharedParamPath = "mySharedParamPath";
        #region Creating a shared parameter file
        System.IO.FileStream fileStream = System.IO.File.Create(mySharedParamPath);
        fileStream.Close();
        #endregion

        #region Getting the definition file
        // set the path of shared parameter file to current Revit
        app.SharedParametersFilename = mySharedParamPath;
        // open the file
        DefinitionFile myDefinitionFile = app.OpenSharedParameterFile();

        #endregion

        // Create a new group in the shared parameters file
        DefinitionGroups myGroupsI = myDefinitionFile.Groups;
        DefinitionGroup myGroupi = myGroupsI.Create("FireCheck Parameters");

        #region Add Parameter


        addInstancParamToElem(uiapp, doc, myGroupi, "Occupant Density (#/m2)", BuiltInCategory.OST_Rooms, true, ParameterType.Number, "Specify the number of occupants per square meter of the area");

        #endregion Add Parameter

        #region Add param to Rooms
        RoomFilter filter = new RoomFilter();
        FilteredElementCollector collector = new FilteredElementCollector(doc);
        IList<Element> rooms = collector.WherePasses(filter).ToElements();

        #endregion

    }



    public void addInstancParamToElem(UIApplication uiapp, Document doc, DefinitionGroup myGroupi, string ParamName, BuiltInCategory BInC, bool modifyBool, ParameterType pType, string tooltip)
    {

        ExternalDefinitionCreationOptions optionI = new ExternalDefinitionCreationOptions(ParamName, pType);
        // let the user to modify the value, only the API
        optionI.UserModifiable = modifyBool;
        // Set tooltip
        optionI.Description = tooltip;
        Definition myDefinition = myGroupi.Definitions.Create(optionI);

        // Create a category set and insert category of door to it
        CategorySet myinstCategories = uiapp.Application.Create.NewCategorySet();
        // Use BuiltInCategory to get category of door
        Category myInstCategory = Category.GetCategory(uiapp.ActiveUIDocument.Document, BInC);

        using (Transaction t = new Transaction(doc, "Add New Parameters"))
        {
            t.Start();


            #region Set param to Element
            myinstCategories.Insert(myInstCategory);

            //Create an instance of InstanceBinding
            InstanceBinding instnceBinding = uiapp.Application.Create.NewInstanceBinding(myinstCategories);

            // Get the BingdingMap of current document.
            BindingMap bindingMap = uiapp.ActiveUIDocument.Document.ParameterBindings;

            // Bind the definitions to the document
            bool typeBindOK = bindingMap.Insert(myDefinition, instnceBinding,
            BuiltInParameterGroup.PG_TEXT);
            #endregion

            t.Commit();

        }

    }


    public static Parameter GetParameter( Element e, Guid guid)
    {
        Parameter parameter = null;
        try
        {
            if (e.get_Parameter(guid) != null) parameter = e.get_Parameter(guid);
            else
            {
                ElementType et = e.Document.GetElement(e.GetTypeId()) as ElementType;
                if (et != null) { parameter = et.get_Parameter(guid); }
                else
                {
                    Material m = e.Document.GetElement(e.GetMaterialIds(false).First()) as Material;
                    parameter = m.get_Parameter(guid);
                }
            }

        }
        catch { }

        return parameter;
    }

}   // End of Class
