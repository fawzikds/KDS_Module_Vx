using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace KDS_Module
{

    #region // create_update_Fam_KDS_CSV
    //    - Separates the large KDS_CSV file into groups per family.
    //    - Copy groups to csv with same family name in dir
    //    - Create Parameters in Family
    //    - Assign Value/Formula to created Parameters
    //    - Save and load family after editing.

    public class create_update_Fam_KDS_CSV : IExternalCommand
    {
        #region  // Setup Revit Revision Dependent csvFilePath
        const string csvFilePath = "Z:\\BIM\\Families\\SupplierCode\\CI_NH_fittings_Charlotte_KDS.csv";    // COntains HPH and List Price for all Elements used by KDS, fittings, Accessories, Finish, etc..
#if (RVT2022)
        const string sharedParameter_dir = "Z:\\BIM\\KDS_TEMPLATE\\2022\\";
        const string sharedParameter_fn = sharedParameter_dir + "KDS_SHARED_PARAMS.txt";            // Shared Parmeters.. this is how we get into the family parameters.  i am still not sure if this is best to use.
#elif (RVT2021)
        const string sharedParameter_dir = "Z:\\BIM\\KDS_TEMPLATE\\2021\\";
        const string sharedParameter_fn = sharedParameter_dir + "KDS_SHARED_PARAMS.txt";
#elif (RVT2020)
        const string sharedParameter_dir = "Z:\\BIM\\KDS_TEMPLATE\\2020\\";
        const string sharedParameter_fn = sharedParameter_dir + "KDS_SHARED_PARAMS.txt";
#elif (RVT2019)
        const string sharedParameter_dir = "Z:\\BIM\\KDS_TEMPLATE\\2019\\";
        const string sharedParameter_fn = sharedParameter_dir + "KDS_SHARED_PARAMS.txt";
#elif (RVT2018)
        const string sharedParameter_dir = "Z:\\BIM\\KDS_TEMPLATE\\2018\\";
        const string sharedParameter_fn = sharedParameter_dir + "KDS_SHARED_PARAMS.txt";
#else
        const string sharedParameter_dir = "Z:\\BIM\\KDS_TEMPLATE\\2022\\";
        const string sharedParameter_fn = sharedParameter_dir + "KDS_SHARED_PARAMS.txt";
#endif
        #endregion  //End Of Setup Revit Revision Dependent csvFilePath

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            #region // Initializations, Filters, Collections //

            UIDocument uidoc = commandData.Application.ActiveUIDocument;

            List<Element> famElem_lst;
            List<FamilySymbol> famSmbl_lst = new List<FamilySymbol>();
            List<string> famNm_lst = new List<string>();
            List<Family> famMain_lst = new List<Family>();
            List<Family> famDstnct_lst = new List<Family>();

            List<BuiltInCategory> builtInCats = new List<BuiltInCategory>();

            builtInCats.Add(BuiltInCategory.OST_PipeFitting);
            //builtInCats.Add(BuiltInCategory.OST_PlumbingFixtures);
            //builtInCats.Add(BuiltInCategory.OST_PipeAccessory);

            ElementMulticategoryFilter filter1 = new ElementMulticategoryFilter(builtInCats);

            famElem_lst = new FilteredElementCollector(uidoc.Document).WherePasses(filter1).WhereElementIsElementType().ToList();

            foreach (Element el in famElem_lst)
            {
                famSmbl_lst.Add(el as FamilySymbol);
            }

            TaskDialog.Show("Execute", "famSmbl_lst.Count: " + famSmbl_lst.Count());


            famMain_lst = famSmbl_lst.Select(fs => fs.Family).ToList();  // Get Family From familySymbols
            famDstnct_lst = famMain_lst //how to find unique families based on family name, not family type, of given categories
                .GroupBy(f => f.Name)
                .Select(g => g.First())
                .ToList();

            TaskDialog.Show("Execute", "famDstnct_lst.Count: " + famDstnct_lst.Count());
            #endregion // Initializations, Filters, Collections //

            CreateAndFill_FamilyDirTree(uidoc, famDstnct_lst);
            return Result.Succeeded;


            #region // Split up CSV file //

            Csv_to_DB(uidoc, famDstnct_lst);   // Convert CVS to a Structured Data
            #endregion

            #region // Import csv in families //
            foreach (Family fam in famDstnct_lst)
            {
                ImportSizeLookUpTable(uidoc, fam);
                Create_sharedParameter_inFamily(uidoc, fam);
            }
            #endregion

            return Result.Succeeded;
        }



        #region   // Create and Fill Families in Dir tree

        public void CreateAndFill_FamilyDirTree(UIDocument uiDoc, List<Family> famDstnct_lst)
        {

            Document actvDoc = uiDoc.Document;

            foreach (Family fam in famDstnct_lst)
            {
                string folderName = null;
                string input = fam.Name;
                int j = input.IndexOf("-");

                if (j >= 0)
                {
                    folderName = input.Substring(0, j);
                }

                string famTypeDir = sharedParameter_dir + "fittings\\";     // z:\BIM\KDS_TEMPLATE\2022\fittings
                string famManfDir = famTypeDir + folderName + "\\";                // z:\BIM\KDS_TEMPLATE\2022\fittings\KDS_Char_CI_NH
                string famDir = famManfDir + fam.Name + "\\";                      // z:\BIM\KDS_TEMPLATE\2022\fittings\KDS_Char_CI_NH\KDS_Char_CI_NH_Coupling\
                //string file = sharedParameter_dir + "fittings\\" + folderName + "\\" + fam.Name + "\\" + fam.Name + ".rfa";   // z:\BIM\KDS_TEMPLATE\2022\fittings\KDS_Char_CI_NH\KDS_Char_CI_NH_Coupling\KDS_Char_CI_NH_Coupling.rfa
                string file = famDir + fam.Name + ".rfa";
               /* TaskDialog.Show("CreateAndFill_FamilyDirTree", "filename = " + fam.Name +
                                           "\n famTypeDir = " + famTypeDir +
                                           "\n famManfDir = " + famManfDir +
                                           "\n famDir = " + famDir +
                                           "\n file = " + file  + "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\");     // The \n makes the dialog box longer, which in turns makes it wider so it can fit the file names.
             */   // If directory does not exist, create it
                //if (!Directory.Exists(famTypeDir)) { Directory.CreateDirectory(famTypeDir); }
               // if (!Directory.Exists(famManfDir)) { Directory.CreateDirectory(famManfDir); }
                if (!Directory.Exists(famDir)) { Directory.CreateDirectory(famDir); }

                try
                {
                    Document currSelFam = actvDoc.EditFamily(fam);

                    SaveAsOptions saveoptions = new SaveAsOptions();
                    saveoptions.OverwriteExistingFile = true;

                    using (Transaction trans = new Transaction(actvDoc, "Load Family"))
                    {
                        currSelFam.SaveAs(file, saveoptions);
                        currSelFam.Close(false);
                    }
                }
                catch (System.Exception ex)
                { TaskDialog.Show("CreateAndFill_FamilyDirTree", file + " Exception: \n" + ex); }

            }   // End For each Family
        }

        #endregion   // End Of Create and Fill Families in Dir tree




        #region  // Convert a CVS File to a Structured Data so we can work with.famDstnct_lst
        public void Csv_to_DB(UIDocument uidoc, List<Family> families)
        {
            foreach (Family fam in families)
            {
                string famName = fam.Name;

                var est_db_hdr = File.ReadLines(csvFilePath)
                    .First(); // For header

                TaskDialog.Show("csv_DB", "hdr:  " + est_db_hdr.ToString());

                var est_db_hdr_1 = est_db_hdr.Skip(1);
                TaskDialog.Show("csv_DB", "hdr:  " + est_db_hdr_1.ToString());

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

                //var est_db_combo = est_db.Where(ds => ds.famName.Equals("KDS_Char_CI_NH_Combo")); //.Select(ds => ds.KDS_MfrList);
                var est_db_fam = est_db.Where(ds => ds.famName.Equals(famName))
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

                int count = est_db_fam.Count();

                string TDS = "count: " + count;

                foreach (var ds in est_db_fam)
                {
                    TDS += "\n ds.KDS_LBR_RATE: " + ds.KDS_LBR_RATE + ":: ds.Size= " + ds.Size + " :: ds.KDS_HPH = " + ds.KDS_HPH + " :: ds.KDS_MfrList = " + ds.KDS_MfrList;
                }

                if (count > 0)
                {
                    TaskDialog.Show("csv2db", "TDS: " + TDS);
                }

                string folderName = null;
                string input = famName;
                int j = input.IndexOf("-");

                if (j >= 0)
                {
                    folderName = input.Substring(0, j);
                }

                string file = sharedParameter_dir + "fittings\\" + folderName + "\\" + fam.Name + "\\" + fam.Name + ".csv";

                sb1 += file + "\n";
                /*string fileToCopy = "c:\\myFolder\\myFile.txt";
                string destinationDirectory = "c:\\myDestinationFolder\\";

                File.Copy(fileToCopy, destinationDirectory + Path.GetFileName(fileToCopy));
*/



                using (var stream = File.CreateText(file))
                {
                    //for (int i = 0; i < reader.Count(); i++)
                    foreach (var ds in est_db_fam)
                    {
                        //string first = reader[i].ToString();
                        //string second = image.ToString();
                        string csvRow = string.Format("{0},{1},{2},{3},{4},{5} ", ds.Size, ds.KDS_HPH, ds.KDS_MfrPart, ds.KDS_MCAA_LBR_RATE, ds.KDS_LBR_RATE, ds.KDS_MfrList);
                        //string.Join(",", line);
                        stream.WriteLine(csvRow);
                    }
                }
                // end foreach famName
                TaskDialog.Show("test", sb1);
            }
        }  // end of CSV_TO_DB
        #endregion  // End Of Convert a CVS File to a Structured Data so we can work with.famDstnct_lst

        public void Create_sharedParameter_inFamily(UIDocument uidoc, Family family)
        {
            //TaskDialog.Show("test", "current family: " + family.Name);
            /// Hard Coded Formulas ///
            string formula1 = "size_lookup(Lookup Table Name, \"F\", 0'  3 1/2\", Main Nominal Diameter)";
            string formula2 = "size_lookup(Lookup Table Name, \"F\", 0'  3 1/2\", Main Nominal Diameter)";
            string formula3 = "size_lookup(Lookup Table Name, \"F\", 0'  3 1/2\", Main Nominal Diameter)";
            string formula4 = "size_lookup(Lookup Table Name, \"F\", 0'  3 1/2\", Main Nominal Diameter)";

            List<string> formulas = new List<string> { formula1, formula2, formula3, formula4 };
            /// Hard Coded Formulas ///



            //  TaskDialog.Show("test", "inside Create_sharedParameter_inFamily");
            Document actvDoc = uidoc.Document;

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


            string folderName = null;
            string input = family.Name;
            int j = input.IndexOf("-");

            if (j >= 0)
            {
                folderName = input.Substring(0, j);
            }

            string file = sharedParameter_dir + "fittings\\" + folderName + "\\" + family.Name + "\\" + family.Name + ".rfa";


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
                        CreateAndSetFormula(ed, uidoc, currSelFam);
                    }
                }

                SaveAsOptions saveoptions = new SaveAsOptions();
                saveoptions.OverwriteExistingFile = true;

                using (Transaction trans = new Transaction(actvDoc, "Load Family"))
                {
                    //currSelFam.LoadFamily(actvDoc, new FamilyLoadOptions());
                    currSelFam.SaveAs(file, saveoptions);
                    currSelFam.Close(false);
                }


            }
            catch (System.Exception ex)
            { TaskDialog.Show("test", "Exception: " + ex); }

        }  // End of create_sharedPamaeter_in_Family

        public void CreateAndSetFormula(ExternalDefinition eDef, UIDocument uidoc, Document currSelFam)
        {
            //TaskDialog.Show("test", "inside CreateSharedParameter");

            //string paramFormula;
            FamilyManager familyManager = currSelFam.FamilyManager;

            using (Transaction y = new Transaction(currSelFam, "Put in parameter"))
            {
                y.Start();
                familyManager.AddParameter(eDef, BuiltInParameterGroup.PG_IDENTITY_DATA, true);
                y.Commit();
            }

            return;
        }

        public void ImportSizeLookUpTable(UIDocument uidoc, Family fam)
        {
            Document actvDoc = uidoc.Document;
            string filePath;
            string sb = null;
            IList<FamilySymbol> fs_lst;
            // IList<Family> f_lst;

            filePath = sharedParameter_dir + "fittings\\"; //subdirectory for fittings only

            fs_lst = Get_PipeFitting_FamilySymbols(actvDoc);
            //TaskDialog.Show("test", "f_lst.count: " + f_lst.Count());

            //foreach (Family fam in f_lst)
            //{
            //    Create_sharedParameter_inFamily(uidoc, fam);
            //}

            string folderName = null;
            string input = fam.Name;
            int j = input.IndexOf("-");

            if (j >= 0)
            {
                folderName = input.Substring(0, j);
            }
            // TaskDialog.Show("test", "folderName: " + folderName);

            Document famDoc = fam.Document;

            FamilySizeTableManager fstm = FamilySizeTableManager.GetFamilySizeTableManager(famDoc, fam.Id);

            using (Transaction importTrans = new Transaction(famDoc))
            {
                importTrans.Start("Start");

                //if (fstm.ExportSizeTable("Really exist name in Lookuptable", AssemblyDirectory + "Really exist name in Lookuptable"+".csv" ))


                string csvFileName = filePath + folderName + "\\" + fam.Name + "\\" + fam.Name + ".csv";
                sb += filePath + folderName + "\\" + fam.Name + "\\" + fam.Name + ".csv\n\n";

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


            //TaskDialog.Show("test", sb);
        }  // end of importSizeLookUpTable

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
    }  //end of Class create_update_Fam_KDS_CSV
    #endregion  // create_update_Fam_KDS_CSV

}  // End of Namespace KDS_Module