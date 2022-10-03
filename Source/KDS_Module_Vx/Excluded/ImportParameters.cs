using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace KDS_Module
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("13AE708F-5403-4432-9F9A-AB440D7BB1D9")]
    public partial class ImportParameters : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            ImportSizeLookUpTable(commandData.Application.ActiveUIDocument);
            Csv_to_DB(commandData.Application.ActiveUIDocument);

            return Result.Succeeded;
        }

        public void Csv_to_DB(UIDocument uidoc)
        {
            List<Element> family_elements;
            List<FamilySymbol> family_symbols = new List<FamilySymbol>();
            List<string> family_names = new List<string>();

            List<BuiltInCategory> builtInCats = new List<BuiltInCategory>();

            builtInCats.Add(BuiltInCategory.OST_PipeFitting);
            //builtInCats.Add(BuiltInCategory.OST_PlumbingFixtures);
            //builtInCats.Add(BuiltInCategory.OST_PipeAccessory);

            ElementMulticategoryFilter filter1 = new ElementMulticategoryFilter(builtInCats);

            family_elements = new FilteredElementCollector(uidoc.Document).WherePasses(filter1).WhereElementIsElementType().ToList();

            foreach (Element ele in family_elements)
            {
                family_symbols.Add(ele as FamilySymbol);
            }

            List<Family> family_mains = family_symbols.Select(fs => fs.Family).ToList();
            List<Family> family_distinct = family_mains //how to find unique families based on family name, not family type, of given categories
                .GroupBy(f => f.Name)
                .Select(g => g.First())
                .ToList();

            foreach (Family fam in family_distinct)
            {
                family_names.Add(fam.Name);
                Create_sharedParameter_inFamily(uidoc, fam);
            }

            string csvFilePath = "Z:\\BIM\\Families\\SupplierCode\\CI_NH_fittings_Charlotte_KDS_test_2.csv";

            var est_db_hdr = File.ReadLines(csvFilePath)
                .First(); // For header

            //TaskDialog.Show("csv_DB", "hdr:  " + est_db_hdr.ToString());

            var est_db_hdr_1 = est_db_hdr.Skip(1);
            //TaskDialog.Show("csv_DB", "hdr:  " + est_db_hdr_1.ToString());

            var est_db = File.ReadAllLines(csvFilePath)
                             .Skip(1) // For header
                             .Select(s => Regex.Match(s, @"^(.*?),(.*?),(.*?),(.*?),(.*?),(.*?),(.*?),(.*?),(.*?),(.*?),(.*?)$"))
                             .Select(data => new
                             {
                                 famName = data.Groups[1].Value,
                                 Size = data.Groups[2].Value,
                                 ND1 = data.Groups[3].Value,
                                 ND2 = data.Groups[4].Value,
                                 ND3 = data.Groups[5].Value,
                                 ND4 = data.Groups[6].Value,
                                 KDS_MfrPart = data.Groups[7].Value,
                                 KDS_HPH = data.Groups[8].Value,
                                 KDS_MfrList = data.Groups[9].Value,
                                 KDS_MCAA_LBR_RATE = data.Groups[10].Value,
                                 KDS_LBR_RATE = data.Groups[11].Value
                             });
            string sb1 = null;

            foreach (string fam in family_names)
            {
                //var est_db_combo = est_db.Where(ds => ds.famName.Equals("KDS_Char_CI_NH_Combo")); //.Select(ds => ds.KDS_MfrList);
                var est_db_combo = est_db.Where(ds => ds.famName.Equals(fam))
                    .Select(ds => new
                    {
                        Size = ds.Size,
                        ND1 = ds.ND1,
                        ND2 = ds.ND2,
                        ND3 = ds.ND3,
                        ND4 = ds.ND4,
                        KDS_MfrPart = ds.KDS_MfrPart,
                        KDS_HPH = ds.KDS_HPH,
                        KDS_MfrList = ds.KDS_MfrList,
                        KDS_MCAA_LBR_RATE = ds.KDS_MCAA_LBR_RATE,
                        KDS_LBR_RATE = ds.KDS_LBR_RATE
                    });



                string TDS = null;



                foreach (var ds in est_db_combo)
                {
                    TDS += "\n ds.KDS_LBR_RATE: " + ds.KDS_LBR_RATE + ":: ds.Size= " + ds.Size + " :: ds.KDS_HPH = " + ds.KDS_HPH + " :: ds.KDS_MfrList = " + ds.KDS_MfrList;
                }

                TaskDialog.Show("csv2db", TDS);

                string folderName = null;
                string input = fam;
                int j = input.IndexOf("-");

                if (j >= 0)
                {
                    folderName = input.Substring(0, j);
                }

                string file = "Z:\\BIM\\KDS_TEMPLATE\\2020\\fittings\\" + folderName + "\\" + fam + "\\" + fam + ".csv";

                sb1 += file + "\n";

                using (var stream = File.CreateText(file))
                {
                    //for (int i = 0; i < reader.Count(); i++)
                    foreach (var ds in est_db_combo)
                    {
                        //string first = reader[i].ToString();
                        //string second = image.ToString();
                        string csvRow = string.Format("{0},{1},{2},{3},{4},{5} ", ds.Size, ds.KDS_HPH, ds.KDS_MfrPart, ds.KDS_MCAA_LBR_RATE, ds.KDS_LBR_RATE, ds.KDS_MfrList);
                        //string.Join(",", line);
                        stream.WriteLine(csvRow);
                    }
                }

                ////////////////////////////////////////////////////////////////////////
            }  // end foreach famName
            TaskDialog.Show("test", sb1);
        }
        // end of CSV_TO_DB    

        public void ImportSizeLookUpTable(UIDocument uidoc)
        {
            Document actvDoc = uidoc.Document;
            string filePath;
            int index = 0;
            string sb = null;
            IList<FamilySymbol> fs_lst;
            // IList<Family> f_lst;

            filePath = "Z:\\BIM\\KDS_TEMPLATE\\2020\\fittings\\"; //subdirectory for fittings only

            fs_lst = Get_PipeFitting_FamilySymbols(actvDoc);
            //TaskDialog.Show("test", "f_lst.count: " + f_lst.Count());

            //foreach (Family fam in f_lst)
            //{
            //    Create_sharedParameter_inFamily(uidoc, fam);
            //}

            foreach (FamilySymbol fs in fs_lst)
            {
                index++;

                string folderName = null;
                string input = fs.FamilyName;
                int j = input.IndexOf("-");

                if (j >= 0)
                {
                    folderName = input.Substring(0, j);
                }
                // TaskDialog.Show("test", "folderName: " + folderName);

                Document famDoc = fs.Family.Document;

                FamilySizeTableManager fstm = FamilySizeTableManager.GetFamilySizeTableManager(famDoc, fs.Family.Id);

                using (Transaction importTrans = new Transaction(famDoc))
                {
                    importTrans.Start("Start");

                    //if (fstm.ExportSizeTable("Really exist name in Lookuptable", AssemblyDirectory + "Really exist name in Lookuptable"+".csv" ))


                    string csvFileName = filePath + folderName + "\\" + fs.FamilyName + "\\" + fs.FamilyName + ".csv";
                    sb += filePath + folderName + "\\" + fs.FamilyName + "\\" + fs.FamilyName + ".csv\n\n";

                    // TaskDialog.Show("importSizeLookup", "\n family name: " + fs.Family.Name + "\n csvFileName: " + csvFileName);

                    FamilySizeTableErrorInfo errorInfo = new FamilySizeTableErrorInfo();

                    bool lookuptableResult = fstm.ImportSizeTable(famDoc, csvFileName, errorInfo);

                    if (lookuptableResult)
                    {
                        // success
                        // TaskDialog.Show("importcsv", "Success for " + fs.FamilyName);
                    }
                    else
                    {
                        // Fail
                        //TaskDialog.Show("importcsv", "Fail" +
                        //                "\n errorInfo.FilePath: " + errorInfo.FilePath +
                        //                "\n errorInfo.GetHashCode: " + errorInfo.GetHashCode());
                    }

                    importTrans.Commit();
                }
            }// end of trans 

            TaskDialog.Show("test", sb);
        }  // end of importSizeLookUpTable



        // Get all pipe fittings families in Document.
        public IList<FamilySymbol> Get_PipeFitting_FamilySymbols(Document actvDoc)
        {
            // WE CANNOT GET FAMILYTPE OUT OF PIPEFITTING CATEGORY  !!!! REVIT DEFECT !!!!
            // SO WE GET INTERSECTION OF PIPEFITTING_CATEGORY LIST WITH FAMILYSYMBOL LIST

            // Find all Pipe Fitting Family types in the document by using category filter
            // Create Pipe Fitting Family Type Filter
            ElementCategoryFilter pipeFitt_fltr = new ElementCategoryFilter(BuiltInCategory.OST_PipeFitting);
            //ElementCategoryFilter pipeFitt_fltr = new ElementCategoryFilter(BuiltInCategory.OST_PlumbingFixtures);

            // Create Pipe Fitting Family Type ollector
            FilteredElementCollector pipeFitt_colc = new FilteredElementCollector(actvDoc);

            // Apply Fillter
            pipeFitt_colc.WherePasses(pipeFitt_fltr).WhereElementIsElementType();
            // copy Collector to List
            IList<Element> pipeFitt_el_lst = pipeFitt_colc.ToList();

            // Find all Pipe Fitting Family types in the document by using category filter
            FilteredElementCollector collector = new FilteredElementCollector(actvDoc);

            // Filter out the familes from all found elements.
            ElementClassFilter filter = new ElementClassFilter(typeof(FamilySymbol));

            // Apply the filter to the elements in the active document 
            collector.WherePasses(filter);

            // Convert queried Families into a list. (safe because ElementClassFilter for Family)
            IList<FamilySymbol> allFamilyTypes_lst = collector.Cast<FamilySymbol>().ToList();


            // This gets the intersection of 2 list of differnt objects based on a common property of these objects.
            // Here i am getting the intersection of FamilySymbols and Elements based of their Id.
            // objA is a Family Symbol, and objB is an Element
            var selectedFamilyTypes_lst = (from objA in allFamilyTypes_lst
                                           join objB in pipeFitt_el_lst on objA.Id equals objB.Id
                                           select objA/*or objB*/).ToList();
            return selectedFamilyTypes_lst;

        }// end of Get_PipeFitting_FamilySymbols    


        public void IndexedList()
        {

            /*
            * String prompt2 = "The Pipe Fitting Families in the current document are:\n";
                       int i =0;

                       foreach (FamilySymbol f in fs_lst)
                       {

                           //Parameter p = f.LookupParameter(eDef_lst[0].Name);

                           prompt2 += i + "- " + f.Family.Name + " :: f.Id = " + f.Id + "\n";
                           i++;
                       }

                       TaskDialog.Show("Revit", prompt2);

              */
        }

        public void Create_sharedParameter_inFamily(UIDocument uidoc, Family family)
        {
            //  TaskDialog.Show("test", "inside Create_sharedParameter_inFamily");
            Document actvDoc = uidoc.Document;
            string sharedParameter_fn;
            sharedParameter_fn = "Z:\\BIM\\KDS_TEMPLATE\\2020\\KDS_SHARED_PARAMS.txt";

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

            string paramFormula;
            FamilyManager familyManager = currSelFam.FamilyManager;

            using (Transaction y = new Transaction(currSelFam, "Put in parameter"))
            {
                y.Start();
                familyManager.AddParameter(eDef, BuiltInParameterGroup.PG_IDENTITY_DATA, true);
                y.Commit();
            }

            //using (Transaction trans = new Transaction(currSelFam, "FP_Formula"))
            //{
            //    trans.Start();

            //        // Check and see if parameter name is KDS_ID_tbl, in which case we need to give it a value of Family name... since this is the csv filenmae as well.
            //        paramFormula = getNewParmValueIfFound(eDef.Name, formula_lst);
            //        if (null != paramFormula)
            //        {
            //            string currParamName = eDef.Name;
            //            if (currParamName == "KDS_ID_tbl")
            //            {
            //                familyManager.SetFormula(eDef, fam.Name);
            //            }
            //            else
            //            {
            //                familyManager.SetFormula(eDef, paramFormula);
            //            }

            //        }//if null != paramFormula
            //}
        
            return;
        }

        public static string getNewParmValueIfFound(string paramName, List<string> paramNames_lst, List<string> paramFormula_lst)
        {
            var paramFormula_zip = paramNames_lst.Zip(paramFormula_lst, (p, f) => new { p, f });
            string nuVal = null;
            foreach (var pf in paramFormula_zip)
            {
                if (pf.p == paramName)
                    nuVal = pf.f;
            }
            return nuVal;

        }  // public static string getNewParmValueIfFound

    }  // end of class
}  // end of namespace
