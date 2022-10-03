using System.Collections.Generic;
 using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace KDS_Module
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("69B6EC93-3643-4B0B-9CCF-1D003180D56A")]
    public class SharedParameters : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();

            string sb = null;

            UIApplication uiapp = commandData.Application;
            Application app = uiapp.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            List<Element> family_elements;
            List<FamilySymbol> family_symbols = new List<FamilySymbol>();

            List<BuiltInCategory> builtInCats = new List<BuiltInCategory>();

            builtInCats.Add(BuiltInCategory.OST_PipeFitting);
            builtInCats.Add(BuiltInCategory.OST_PlumbingFixtures);
            builtInCats.Add(BuiltInCategory.OST_PipeAccessory);

            ElementMulticategoryFilter filter1 = new ElementMulticategoryFilter(builtInCats);

            family_elements = new FilteredElementCollector(uidoc.Document).WherePasses(filter1).WhereElementIsElementType().ToList();

            foreach (Element ele in family_elements)
            {
                family_symbols.Add(ele as FamilySymbol);
            }

            var family_mains = family_symbols.Select(fs => fs.Family).ToList();
            var family_distinct = family_mains //how to find unique families based on family name, not family type, of given categories
                .GroupBy(f => f.Name)
                .Select(g => g.First())
                .ToList();

            foreach (var var in family_distinct)
            {
                sb += "\nfamily: " + var.Name;
                Create_sharedParameter_inFamily(uidoc, var);
            }

            watch.Stop();
            var elapsedTime = watch.ElapsedMilliseconds;

            TaskDialog.Show("test", sb + "\nTime Elapsed: " + elapsedTime / 1000.0);

            return Result.Succeeded;
        }
        public void Create_sharedParameter_inFamily(UIDocument uidoc, Family family)
        {
            //  TaskDialog.Show("test", "inside Create_sharedParameter_inFamily");
            Document actvDoc = uidoc.Document;
            string sharedParameter_fn;
            sharedParameter_fn = "Z:\\BIM\\KDS_TEMPLATE\\KDS_SHARED_PARAMS.txt";

            // Set Filename for SharedParameters
            uidoc.Application.Application.SharedParametersFilename = sharedParameter_fn;
            DefinitionFile defFile = uidoc.Application.Application.OpenSharedParameterFile();

            // Get List of Groups
            DefinitionGroups myGroups = defFile.Groups;
            IList<DefinitionGroup> myGroup_lst = myGroups.ToList();


            // Get Group Definitions.  Definitions an be: 
            //   InternalDefinitions that reside inside a Revit Document, OR
            //   ExternalDefinintions that reside in an external File...such as shared Parameters.
            Definitions myDefinitions = myGroup_lst[0].Definitions;
            List<Definition> Def_lst = myDefinitions.ToList();   //    get_Item("PRL_WIDTH") as ExternalDefinition;

            //  Cast Definitions to ExternalDefinitions..... How did myGroup become a shared parameter. 
            List<ExternalDefinition> eDef_lst = Def_lst.Select(df => df as ExternalDefinition).ToList();

            //TaskDialog.Show("create_sharedPamaeter_in_Family", " Def_lst.Count:   " + Def_lst.Count +
            //                "\n eDef_lst[0].Name:  " + eDef_lst[0].Name +
            //                "\n eDef_lst[1].Name:  " + eDef_lst[1].Name +
            //                "\n eDef_lst[2].Name:  " + eDef_lst[2].Name +
            //                "\n eDef_lst[3].Name:  " + eDef_lst[3].Name +
            //                "\n eDef_lst[4].Name:  " + eDef_lst[4].Name +
            //                "\n eDef_lst[5].Name:  " + eDef_lst[5].Name
            //               );

            //TaskDialog.Show("test", "family: " + family.Name);

            bool found = false;
            try
            {
                Document currSelFam = actvDoc.EditFamily(family);

                FamilyManager familyMngr = currSelFam.FamilyManager;

                IList<FamilyParameter> currFamParam_lst = familyMngr.GetParameters();

                foreach (ExternalDefinition ed in eDef_lst)
                {
                    foreach (FamilyParameter a in currFamParam_lst)
                    {
                        if (a.Definition.Name == ed.Name)
                        {
                            found = true;
                        }
                        else
                        {
                            continue;
                        }
                    }

                    if (found == false)
                    {
                        CreateSharedParameter(ed, uidoc, currSelFam);
                    }
                }
                using (Transaction trans = new Transaction(actvDoc, "Load Family"))
                    currSelFam.LoadFamily(actvDoc, new FamilyLoadOptions());


            }
            catch
            { TaskDialog.Show("test", "Family.Name: " + family.Name); }
            return;
        }  // End of create_sharedPamaeter_in_Family

        public void CreateSharedParameter(ExternalDefinition eDef, UIDocument uidoc, Document currSelFam)
        {
            //TaskDialog.Show("test", "inside CreateSharedParameter");

            FamilyManager familyManager = currSelFam.FamilyManager;

            using (Transaction y = new Transaction(currSelFam, "Put in parameter"))
            {
                y.Start();
                familyManager.AddParameter(eDef, BuiltInParameterGroup.PG_IDENTITY_DATA, true);
                y.Commit();
            }

            return;
        }

        // Get all pipe fittings families in Document.
        public IList<FamilySymbol> get_PipeFitting_FamilySymbols(Document actvDoc)
        {
            // WE CANNOT GET FAMILYTPE OUT OF PIPEFITTING CATEGORY  !!!! REVIT DEFECT !!!!
            // SO WE GET INTERSECTION OF PIPEFITTING_CATEGORY LIST WITH FAMILYSYMBOL LIST

            // Find all Pipe Fitting Family types in the document by using category filter
            // Create Pipe Fitting Family Type Filter
            ElementCategoryFilter pipeFitt_fltr = new ElementCategoryFilter(BuiltInCategory.OST_PipeFitting);

            // Create Pipe Fitting Family Type ollector
            FilteredElementCollector pipeFitt_colc = new FilteredElementCollector(actvDoc);

            // Apply Fillter
            pipeFitt_colc.WherePasses(pipeFitt_fltr).WhereElementIsElementType();
            // copy Collector to List
            IList<Element> pipeFitt_el_lst = pipeFitt_colc.ToList<Element>();

            // Find all Pipe Fitting Family types in the document by using category filter
            FilteredElementCollector collector = new FilteredElementCollector(actvDoc);

            // Filter out the familes from all found elements.
            ElementClassFilter filter = new ElementClassFilter(typeof(FamilySymbol));

            // Apply the filter to the elements in the active document 
            collector.WherePasses(filter);

            // Convert queried Families into a list. (safe because ElementClassFilter for Family)
            IList<FamilySymbol> allFamilyTypes_lst = collector.Cast<FamilySymbol>().ToList<FamilySymbol>();


            // This gets the intersection of 2 list of differnt objects based on a common property of these objects.
            // Here i am getting the intersection of FamilySymbols and Elements based of their Id.
            // objA is a Family Symbol, and objB is an Element
            var selectedFamilyTypes_lst = (from objA in allFamilyTypes_lst
                                           join objB in pipeFitt_el_lst on objA.Id equals objB.Id
                                           select objA/*or objB*/).ToList();
            return selectedFamilyTypes_lst;

        }// end of Get_PipeFitting_FamilySymbols



    } // end of Class SharedParameters

    class FamilyLoadOptions : IFamilyLoadOptions
    {
        bool IFamilyLoadOptions.OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
        {
            overwriteParameterValues = true;
            return true;
        }

        bool IFamilyLoadOptions.OnSharedFamilyFound(Family sharedFamily, bool familyInUse, out FamilySource source, out bool overwriteParameterValues)
        {
            source = FamilySource.Family;
            overwriteParameterValues = false;
            return true;
        }
    }  // end of Class familyLoadOptions

} // KDS_Tools namespace