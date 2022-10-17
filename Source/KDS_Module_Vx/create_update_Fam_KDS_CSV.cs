using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

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
        //const string csvFilePath_const = "Z:\\BIM\\Families\\SupplierCode\\CI_NH_fittings_Charlotte_KDS.csv";    // COntains HPH and List Price for all Elements used by KDS, fittings, Accessories, Finish, etc..
        //const string csvFilePath_const = "Z:\\BIM\\Families\\SupplierCode\\KDS_CI_NH_BLK_MLBL.csv";    // COntains HPH and List Price for all Elements used by KDS, fittings, Accessories, Finish, etc..
        //const string csvFilePath_const = get_PriceDB_fileName();

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
            List<Family> DstnctFam_lst = new List<Family>();
            List<bicFam_strct> bicFam_lst = new List<bicFam_strct>();   // holds list of all families per bic
            List<bicFam_strct> bicDstnctFam_lst = new List<bicFam_strct>();    // Holds lists of distinct famlies per bic
            string csvFilePath_const = get_PriceDB_fileName();
            #endregion // Initializations, Filters, Collections //

            #region   // Add the Category of Families to use (Pipe Fittings, Accessories, Fixtures...)
            List<BuiltInCategory> bic_lst = new List<BuiltInCategory>();

            //bic_lst.Add(BuiltInCategory.OST_PipeFitting);
            bic_lst.Add(BuiltInCategory.OST_PlumbingFixtures);
            //bic_lst.Add(BuiltInCategory.OST_PipeAccessory);
            #endregion  // End Of Add the Category of Families to use (Pipe Fittings, Accessories, Fixtures...)

            #region // Loop Thru every Category of Families and get all of its Elements in the actvDocument
            // Then store in bicFam structure so we can access all families by their Category.
            string bicfam_tds = " Family Count per BIC\n";
            foreach (BuiltInCategory bic in bic_lst)
            {
                famElem_lst = new FilteredElementCollector(uidoc.Document).OfCategory(bic).WhereElementIsElementType().ToList();
                famSmbl_lst = new FilteredElementCollector(uidoc.Document).OfCategory(bic).WhereElementIsElementType().Select(df => df as FamilySymbol).ToList();
                famMain_lst = famSmbl_lst.Select(fs => fs.Family).ToList();  // Get Family From familySymbols
                // This get a unique list of all families per Category.
                DstnctFam_lst = famMain_lst.GroupBy(f => f.Name).Select(g => g.First()).ToList();    //how to find unique families based on family name, not family type, of given categories

                // Load List for families per bic
                bicFam_strct bicFam = new bicFam_strct();
                bicFam.bic_str = bic.ToString();
                bicFam.fam_lst = famMain_lst;  //DstnctFam_lst;
                bicFam_lst.Add(bicFam);

                // Load List for Distinct family nmaes per bic
                bicFam_strct bicDstnctFam = new bicFam_strct();
                bicDstnctFam.bic_str = bic.ToString();
                bicDstnctFam.fam_lst = DstnctFam_lst;
                bicDstnctFam_lst.Add(bicDstnctFam);

                bicfam_tds += "\n BIC: " + bicFam.bic_str + ":: Count = " + bicFam.fam_lst.Count();
                //TaskDialog.Show("Execute", "BIC: " + bicFam.famTyp + ":: Count = " + bicFam.fam_lst.Count()+ "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");
            }
            //TaskDialog.Show("Execute", bicfam_tds + "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");

            #endregion // End Of Loop Thru every Category of Families and get all of its Elements in the actvDocument

            #region   // Loop thru every Category and Create the Dir Tree if it does not exist  including .rfa and its associated but empty .csv
            if (bicFam_lst.Count() > 0)
            {
                string ftf_tds = "famSmbl_lst.Count: " + bicFam_lst.Count();
                foreach (bicFam_strct bicfam in bicFam_lst)
                {
                    ftf_tds += "\n Family BIC: " + bicfam.bic_str + ":: Count: " + bicfam.fam_lst.Count;
                    if (bicfam.fam_lst.Count > 0)
                    {
                        CreateAndFill_FamilyDirTree(uidoc, bicfam.bic_str, csvFilePath_const, bicfam.fam_lst);
                    }
                }  // End of for each FamTyFam ftf
                //TaskDialog.Show("Execute", ftf_tds + "\n\n\n\n\n\n\n\n\n\n\n\n");
            }// End if bicFam_lst is >0
            #endregion   // End OF Loop thru every Category and Create the Dir Tree if it does not exist

            #region    // Loop thru every Category and Create CSV from est_DB file for every Distinct Family 
            show_DstnctFam(bicDstnctFam_lst); // Loop thru every Category and show Distinct Families  -- Debug Only

            if (bicDstnctFam_lst.Count() > 0)
            {
                string ftf_tds = "bicDstnctFam.Count: " + bicFam_lst.Count();
                foreach (bicFam_strct bicDstnctFam in bicDstnctFam_lst)
                {
                    if (bicDstnctFam.fam_lst.Count > 0)
                    {
                        Csv_to_DB(uidoc, csvFilePath_const, bicDstnctFam.bic_str, bicDstnctFam.fam_lst);   // Convert CVS to a Structured Data
                    }
                }  // End of for each FamTyFam ftf

            }// End if bicFam_lst is >0
            #endregion    // End Of Loop thru every Category and Create CSV from est_DB file for every Distinct Family 

            #region // Import csv into their respective families //

            string imprt_tds = "";
            if (bicDstnctFam_lst.Count() > 0)
            {
                foreach (bicFam_strct bicDstnctFam in bicDstnctFam_lst)
                {
                    // Reset string per BIC
                    imprt_tds = bicDstnctFam.bic_str + "\n";
                    TaskDialog.Show("Execute", " Getting Ready to Import csvFiles for: " + bicDstnctFam.bic_str);
                    if (bicDstnctFam.fam_lst.Count > 0)
                    {
                        imprt_tds = "bicDstnctFam.fam_lst.Count: " + bicDstnctFam.fam_lst.Count;
                        int ftf_cnt = 0;

                        foreach (Family DstnctFam in bicDstnctFam.fam_lst)
                        {
                            ftf_cnt++;
                            imprt_tds += "\n " + ftf_cnt + "- " + DstnctFam.Name + ":: Errors: " + ImportSizeLookUpTable(uidoc, bicDstnctFam.bic_str, DstnctFam);

                            //Create_sharedParameter_inFamily(uidoc, bicDstnctFam.bic, DstnctFam);
                        }  // foreach bicDstnctFam
                        TaskDialog.Show("Execute", imprt_tds + "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");
                    }  // End Of if  bicDstnctFam.fam_lst.Count > 0
                }  // End Of foreach bicDstnctFam in bicDstnctFam_lst
            }  // End Of if bicDstnctFam_lst.Count() > 0

            #endregion

            return Result.Succeeded;
        }


        #region   // Create and Fill Families in Dir tree  with the .rfa and the .csv (csv will be empty now)
        public void CreateAndFill_FamilyDirTree(UIDocument uiDoc, String famBIC_Nm, string csvFilePath_const, List<Family> DstnctFam_lst)
        {
            Autodesk.Revit.DB.Document actvDoc = uiDoc.Document;

            foreach (Family dstnctFam in DstnctFam_lst)
            {
                string folderName = null;
                string input = dstnctFam.Name;

                int j = input.IndexOf("-");

                if (j >= 0)
                {
                    folderName = input.Substring(0, j);
                }

                //string famTypeDir = sharedParameter_dir + BuiltInCategory.OST_PipeFitting.ToString() + "\\";     // z:\BIM\KDS_TEMPLATE\2022\fittings
                string famTypeDir = sharedParameter_dir + famBIC_Nm + "\\";     // z:\BIM\KDS_TEMPLATE\2022\fittings
                string famManfDir = famTypeDir + folderName + "\\";                // z:\BIM\KDS_TEMPLATE\2022\fittings\KDS_Char_CI_NH
                string famDir = famManfDir + dstnctFam.Name + "\\";                      // z:\BIM\KDS_TEMPLATE\2022\fittings\KDS_Char_CI_NH\KDS_Char_CI_NH_Coupling\
                                                                                         //string file = sharedParameter_dir + "fittings\\" + folderName + "\\" + fam.Name + "\\" + fam.Name + ".rfa";   // z:\BIM\KDS_TEMPLATE\2022\fittings\KDS_Char_CI_NH\KDS_Char_CI_NH_Coupling\KDS_Char_CI_NH_Coupling.rfa

                string DstnctFamFname = get_FileNames(sharedParameter_dir, famBIC_Nm, dstnctFam.Name, ".rfa");


                /* DEBUG  TaskDialog.Show("CreateAndFill_FamilyDirTree", "filename = " + fam.Name +
                                            "\n famTypeDir = " + famTypeDir +
                                            "\n famManfDir = " + famManfDir +
                                            "\n famDir = " + famDir +
                                            "\n file = " + file + "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");     // The \n makes the dialog box longer, which in turns makes it wider so it can fit the file names.
 */
                //If directory does not exist, create it
                if (!Directory.Exists(famDir)) { Directory.CreateDirectory(famDir); }

                // if file Does not exists or if it is older than curren... recreate it... this is now Always Overwrite 
                if (!System.IO.File.Exists(DstnctFamFname) || (System.IO.File.GetCreationTime(DstnctFamFname) < DateTime.UtcNow.AddMonths(-1)))
                {
                    try
                    {
                        Autodesk.Revit.DB.Document currSelFam = actvDoc.EditFamily(dstnctFam);

                        SaveAsOptions saveoptions = new SaveAsOptions();
                        saveoptions.OverwriteExistingFile = true;

                        using (Transaction saveFam_trx = new Transaction(actvDoc, "Save Family"))
                        {
                            saveFam_trx.Start();
                            currSelFam.SaveAs(DstnctFamFname, saveoptions);
                            currSelFam.Close(false);
                            saveFam_trx.Commit();
                            saveFam_trx.Dispose();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        TaskDialog.Show("CreateAndFill_FamilyDirTree", DstnctFamFname + " Exception: \n" + ex);
                    }
                }  // End of if File Exists or older
                #region  // Create an Empty csv for each family lookup table file.  --- Uneccessary
                //i don't need this since i will create it later as i am filling them with content from the KDS  est csv file.
                // However, since i am still in testing, i need to have a file for every family so i can test the importLookuptable function.
                // But this will not work for testing adding other arameters since the filles would not have any content.... unless if i add the header only ???
                //Add CSV Files as Well
                //if file Does not exists or if it is older than curren... recreate it
                est_data_class est_db_hdr = est_data_class.FromCsv(System.IO.File.ReadAllLines(csvFilePath_const).First());
                string csvRow = "";
                foreach (Family DstnctFam in DstnctFam_lst)
                {
                    string csvFile = get_FileNames(sharedParameter_dir, famBIC_Nm, DstnctFam.Name, ".csv");
                    System.IO.Directory.CreateDirectory(Path.GetDirectoryName(csvFile));
                    using (

                        var stream = System.IO.File.CreateText(csvFile))     // OPens or Creates File if it does not exist. if It exists it will overwrite its content
                    {
                        // Write Headers first
                        csvRow = est_db_hdr.ToCsvString_excld("Size");
                        stream.WriteLine(csvRow);
                    }  // End Of foreach est_db_fam
                }  // End Of Using
                #endregion  // END Of Create an Empty csv for each family lookup table file.  --- Uneccessary
            }   // End For each Family
        }  // End of CreateAndFill_FamilyDirTree
        #endregion   // End Of Create and Fill Families in Dir tree



        #region  // Convert a CVS File to a Structured Data so we can work with.DstnctFam_lst
        // we have to use the Distinct Family list, (DstnctFam_lst) since we will have multiple lines that
        // represents different fitting sizes for each family that will need to go into the same csv file.
        public void Csv_to_DB(UIDocument uidoc, string csvFilePath_const, string famBIC_Nm, List<Family> DstnctFam_lst)
        {

            #region // test input lst   -- Debug Only
            /*          string ftf_tds = "\n Family BIC: " + famBIC_Nm + ":: Count: " + DstnctFam_lst.Count;
                      if (DstnctFam_lst.Count > 0)
                      {
                          int ftf_bic_int = 0;
                          foreach (Family dstnctfam in DstnctFam_lst)
                          {
                              ftf_bic_int++;
                              ftf_tds += "\n" + ftf_bic_int + "- " + dstnctfam.Name;
                          }
                          TaskDialog.Show("Csv_to_DB", ftf_tds + "\n\n\n\n\n\n\n\n\n\n\n\n");
                      }
                      else { TaskDialog.Show("Csv_to_DB", "Distinct Family List Count is: " + DstnctFam_lst.Count); }
          */
            #endregion // End Of test input list

            #region  // Get Data from the KDS Estimation CSV File.
            // I seperated  getting the header and the data as 2 IO transactions since i am not
            // sure if i loaded the file once into a var, then i split it into a hdr and data,
            // i will have the file in memory, twice.
            // Get all Data from KDS est Files .. This will include Headers.
            est_data_class est_db_hdr = new est_data_class();
            List<est_data_class> est_data_class_lst = new List<est_data_class>();
            // File may be locked by another process use try.
            try
            {
                // Get the First line from the KDS Est Data File as Column Headers

                est_db_hdr = est_data_class.FromCsv(System.IO.File.ReadAllLines(csvFilePath_const).First());
                //TaskDialog.Show("csv_DB", "BIC : " + famBIC_Nm + "\n  - Formatted Header: \n\n " + est_db_hdr.ToFormatString() + "\n\n - CSV Format: " + est_db_hdr.ToCsvString_wfname());

                est_data_class_lst = System.IO.File.ReadAllLines(csvFilePath_const)
                                                        .Skip(1)
                                                        .Select(v => est_data_class.FromCsv(v))
                                                        .ToList();
                //TaskDialog.Show("csv_DB", "BIC : " + famBIC_Nm + "\n - Formatted Data:  \n\n" + est_data_class_lst[2].ToFormatString() + "\n\n - CSV Format: " + est_data_class_lst[2].ToCsvString_wfname());
                //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! The Contents seems to be the same regardless of BIC   !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            }
            catch (Exception ex)
            {
                TaskDialog.Show("csv_DB", " Exception in accessing KDS Estimation File:" + ex);
            }

            string tds_edc = "BIC: " + famBIC_Nm + ":: CSV File toList Count: " + est_data_class_lst.Count;
            int edc_cnt = 0;
            foreach (est_data_class est_data_class in est_data_class_lst)
            {
                edc_cnt++;
                tds_edc += "\n " + edc_cnt + "- Formatted Data: " + est_data_class.ToCsvString_wfname();
            }
            //TaskDialog.Show("Csv_to_DB", tds_edc + "\n\n\n\n\n\n\n\n\n");
            #endregion  // Get Data from the KDS Estimation CSV File.

            #region  // Loop Thru every Distinct Family and get its associated Rows from the Data var (est_db_data); Then Save these rows to a csv file.
            string dstnctFam_summary = " Distinct Family Summary for Category: " + famBIC_Nm + "\n";
            int dstnctFam_summary_cnt = 0;

            #region // Loop Thru every Distinct Family and get its associated Rows from the Data var (est_db_data);
            foreach (Family DstnctFam in DstnctFam_lst)
            {
                string DstnctFam_nm = DstnctFam.Name;
                //TaskDialog.Show("Csv_to_DB", "BIC: " + famBIC_Nm + ":: Currrent Distinct Family Name is: " + DstnctFam_nm + "\n\n\n\n\n\n\n\n\n\n");

                List<est_data_class> est_db_DstnctFam_lst = est_data_class_lst.Where(dc => dc.famName.Equals(DstnctFam_nm)).ToList<est_data_class>();    // Holds the rows per Family to save a csv

                if (est_db_DstnctFam_lst.Count > 0)
                {
                    dstnctFam_summary_cnt++;
                    dstnctFam_summary += "\n " + dstnctFam_summary_cnt + "- Currrent Distinct Family " + DstnctFam_nm + " count: " + est_db_DstnctFam_lst.Count;

                    //dstnctFam_summary += "\n Currrent Distinct Family " + DstnctFam_nm + " count: " + est_db_DstnctFam_lst.Count;
                    //TaskDialog.Show("Csv_to_DB", "BIC: " + famBIC_Nm + ":: Currrent Distinct Family " + DstnctFam_nm + " count: " + est_db_DstnctFam_lst.Count + "\n\n\n\n\n\n\n\n\n\n");
                    string TDS = "BIC: " + famBIC_Nm + "\n Distinct Family Count:" + est_db_DstnctFam_lst.Count() + "\n     Only One Family Name Should Appear Here    \n";
                    int row_cnt = 0;

                    foreach (var est_db_fam in est_db_DstnctFam_lst)
                    {
                        row_cnt++;
                        TDS += "\n" + est_db_fam.ToCsvString_wfname();
                    }

                    //TaskDialog.Show("Csv_to_DB", TDS + "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");

                    #endregion // End Of Loop Thru every Distinct Family and get its associated Rows from the Data var (est_db_data);

                    //Create CSV File Name

                    string DstnctFamFname = get_FileNames(sharedParameter_dir, famBIC_Nm, DstnctFam.Name, ".csv");

                    dstnctFam_summary += " -- CSV File Name:" + DstnctFamFname;

                    #region  // Save these rows to a csv file // OPens or Creates File if it does not exist. if It exists it will overwrite its content
                    using (
                        var stream = System.IO.File.CreateText(DstnctFamFname))     // OPens or Creates File if it does not exist. if It exists it will overwrite its content
                    {
                        string csvRow = "";
                        // Write Headers first
                        //csvRow = string.Format("{0},{1},{2},{3},{4},{5} ", est_db_hdr.Size, est_db_hdr.KDS_HPH, est_db_hdr.KDS_MfrPart, est_db_hdr.KDS_MCAA_LBR_RATE, est_db_hdr.KDS_LBR_RATE, est_db_hdr.KDS_MfrList);
                        csvRow = est_db_hdr.ToCsvString_excld("Size");
                        stream.WriteLine(csvRow);

                        // Write the Rest of the data
                        foreach (est_data_class ds in est_db_DstnctFam_lst)
                        {
                            //csvRow = string.Format("{0},{1},{2},{3},{4},{5} ", ds.Size, ds.KDS_HPH, ds.KDS_MfrPart, ds.KDS_MCAA_LBR_RATE, ds.KDS_LBR_RATE, ds.KDS_MfrList);
                            csvRow = ds.ToCsvString();
                            stream.WriteLine(csvRow);
                        }  // End Of foreach est_db_fam
                    }  // End Of Using


                    //TaskDialog.Show("CSV_TO_DB", sb1);

                    #endregion  // End Of Save Found Famiy rows to a csv file.
                }  // end foreach famName

            }
            //TaskDialog.Show("Csv_to_DB", "TDS: " + dstnctFam_summary + "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");
            #endregion  // Loop Thru every Distinct Family and get its associated Rows from the Data var (est_db_data); Then Save these rows to a csv file.
        }  // end of CSV_TO_DB
        #endregion  // End Of Convert a CVS File to a Structured Data so we can work with.DstnctFam_lst

        #region  // Create_sharedParameter_inFamily 
        public void Create_sharedParameter_inFamily(UIDocument uidoc, string famBIC_Nm, Family family)
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
            Autodesk.Revit.DB.Document actvDoc = uidoc.Document;

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


            string file = get_FileNames(sharedParameter_dir, famBIC_Nm, family.Name, ".rfa");

            bool found = false;
            try
            {
                Autodesk.Revit.DB.Document currSelFam = actvDoc.EditFamily(family);

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
                        CreateAndSetFormula(ed, currSelFam);
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
        #endregion  // End Of Create_sharedParameter_inFamily 

        #region  // CreateAndSetFormula 
        // 
        public void CreateAndSetFormula(ExternalDefinition eDef, Autodesk.Revit.DB.Document currSelFam)
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
        #endregion  // End Of CreateAndSetFormula 




        #region // ImportSizeLookUpTable to Families
        // Imports the CSV lookup table file into a family.
        public string ImportSizeLookUpTable(UIDocument uidoc, string bic_nm, Family DstnctFam)
        {
            // Get RFA and CSV File Name
            string DstnctFamPath = get_FileNames(sharedParameter_dir, bic_nm, DstnctFam.Name, ".rfa");
            string csvFilePath = get_FileNames(sharedParameter_dir, bic_nm, DstnctFam.Name, ".csv");
            //TaskDialog.Show("ImportSizeLookUpTable", "\n Document File Name : " + DstnctFamPath + "\n With csvFilePath: " + csvFilePath);

            try
            {
                Autodesk.Revit.DB.Document actvDoc = uidoc.Document;
                string importError = "";
                // Get the Family's Size and Table Manager


                Autodesk.Revit.DB.Document DstnctFamEdt = actvDoc.EditFamily(DstnctFam);
                Autodesk.Revit.DB.FamilySizeTableManager fstm = Autodesk.Revit.DB.FamilySizeTableManager.GetFamilySizeTableManager(DstnctFamEdt, DstnctFamEdt.OwnerFamily.Id);
                if (fstm.IsValidObject)
                {
                    // Import the csv file into Family.
                    Transaction importTrans = new Transaction(DstnctFam.Document, "Importing csv File into Family ");

                    //TaskDialog.Show("ImportSizeLookUpTable", "Starting with: " + DstnctFam.Name);
                    importTrans.Start("Start");

                    // Create Error Info  for ImportSizeTable
                    FamilySizeTableErrorInfo errorInfo = new FamilySizeTableErrorInfo();
                    //TaskDialog.Show("ImportSizeLookUpTable", "Created errorInfo ");

                    Autodesk.Revit.DB.FamilySizeTableErrorInfo ImportErrorInfo = new Autodesk.Revit.DB.FamilySizeTableErrorInfo();
                    fstm.ImportSizeTable(DstnctFamEdt, csvFilePath, ImportErrorInfo);
                    //TaskDialog.Show("ImportSizeLookUpTable", "after  ImportSizeTable ");
                    // End transaction
                    importTrans.Commit();
                    //TaskDialog.Show("ImportSizeLookUpTable", "after  Commit ");

                    DstnctFamEdt.LoadFamily(actvDoc, new familyLoadOptions()); // here the update begins
                    //TaskDialog.Show("ImportSizeLookUpTable", "after  LoadFamily ");

                    SaveAsOptions saveoptions = new SaveAsOptions();
                    saveoptions.OverwriteExistingFile = true;

                    try
                    {
                        DstnctFamEdt.SaveAs(DstnctFamPath, saveoptions);  // Error is here
                        //TaskDialog.Show("ImportSizeLookUpTable", "after  SaveAs ");

                    }
                    catch (Exception ex)
                    {
                        TaskDialog.Show("ImportSizeLookUpTable", "Error SaveAs after Loading family. \n Exception: " + ex);
                    }

                    try
                    {
                        DstnctFamEdt.Close(false);
                        //TaskDialog.Show("ImportSizeLookUpTable", "after  Close ");
                    }
                    catch (Exception ex)
                    {

                        TaskDialog.Show("ImportSizeLookUpTable", "Error Close after Loading family. \n Exception: " + ex);
                    }
                }
                return importError;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("ImportSizeLookUpTable", " DstnctFam.Name: " + DstnctFam.Name + "\n With CSV: " + csvFilePath + "\n - lookuptableResult excepton: " + ex);
                throw;
            }

        }  // end of importSizeLookUpTable
        #endregion // End Of ImportSizeLookUpTable to Families
        public IList<FamilySymbol> Get_PipeFitting_FamilySymbols(Autodesk.Revit.DB.Document actvDoc)
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


        #region // Get file Names
        public string get_FileNames(string sharedParameter_dir, string bic_nm, string famName, string ext)
        {
            string folderName = null;

            int j = famName.IndexOf("-");

            if (j >= 0)
            {
                folderName = famName.Substring(0, j);
            }

            string famTypeDir = sharedParameter_dir + bic_nm + "\\";              // z:\BIM\KDS_TEMPLATE\2022\fittings
            string famManfDir = famTypeDir + folderName + "\\";                // z:\BIM\KDS_TEMPLATE\2022\fittings\KDS_Char_CI_NH
            string famDir = famManfDir + famName + "\\";                       // z:\BIM\KDS_TEMPLATE\2022\fittings\KDS_Char_CI_NH\KDS_Char_CI_NH_Coupling\

            string file = famDir + famName + ext;                        // z:\BIM\KDS_TEMPLATE\2022\fittings\KDS_Char_CI_NH\KDS_Char_CI_NH_Coupling\KDS_Char_CI_NH_Coupling.???

            return file;
        }  // End Of getFileNames
        #endregion // Enf Of Get FIle Names


        #region  // Get Pricing Database filename
        public string get_PriceDB_fileName()
        {
            /*string selectedFileName = "";
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            //openFileDialog1.Filter = "Database files (*.mdb, *.accdb)|*.mdb;*.accdb";
            openFileDialog1.Filter = "Database files (*.csv, *.csv)|*.csv;*.csv";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                selectedFileName = openFileDialog1.FileName;
            }
            else { selectedFileName = "Z:\\BIM\\Families\\SupplierCode\\KDS_CI_NH_BLK_MLBL.csv"; }
            openFileDialog1.Dispose();
            return selectedFileName;*/
            return "Z:\\BIM\\Families\\SupplierCode\\KDS_CI_NH_BLK_MLBL.csv";
        }
        #endregion  // End Of Get Pricing Database filename

        #region // bicFam_strct
        public struct bicFam_strct
        {
            public string bic_str;
            public List<Family> fam_lst;

            public bicFam_strct(string bic_str, List<Family> fam_lst)
            {
                this.bic_str = bic_str;
                this.fam_lst = fam_lst;

            }

        } // End of bicFam_strct
        #endregion  // End of bicFam_strct


        #region   // Loop thru every Category and show Distinct Families   -- For Debug Only
        public void show_DstnctFam(List<bicFam_strct> bicDstnctFam_lst)
        {
            if (bicDstnctFam_lst.Count() > 0)
            {
                //string ftf_tds = "bicDstnctFam.Count: " + bicFam_lst.Count();
                string ftf_tds = "bicDstnctFam.Count: " + bicDstnctFam_lst.Count();
                foreach (bicFam_strct bicDstnctFam in bicDstnctFam_lst)
                {
                    ftf_tds = "\n Family BIC: " + bicDstnctFam.bic_str + ":: Count: " + bicDstnctFam.fam_lst.Count;
                    if (bicDstnctFam.fam_lst.Count > 0)
                    {
                        int ftf_bic_int = 0;
                        foreach (Family dstnctfam in bicDstnctFam.fam_lst)
                        {
                            ftf_bic_int++;
                            ftf_tds += "\n" + ftf_bic_int + "- " + dstnctfam.Name;
                        }
                        //TaskDialog.Show("Execute", ftf_tds + "\n\n\n\n\n\n\n\n\n\n\n\n");
                    }
                }  // End of for each FamTyFam ftf
            }// End if bicFam_lst is >0
        }   // End OF show_DstnctFam
        #endregion   // End OF Loop thru every Category and Create the Dir Tree if it does not exist












    }  //end of Class create_update_Fam_KDS_CSV
    #endregion  // create_update_Fam_KDS_CSV
    //======================================== EST_DATA_CLASS ===========================================================//

    #region  //est_data_class
    class est_data_class
    {

        public string Size { get; set; }
        public string famName { get; set; }
        public string ND1 { get; set; }
        public string ND2 { get; set; }
        public string ND3 { get; set; }
        public string ND4 { get; set; }
        public string KDS_MfrPart { get; set; }
        public string KDS_HPH { get; set; }
        public string KDS_MfrList { get; set; }
        public string KDS_MCAA_LBR_RATE { get; set; }
        public string KDS_LBR_RATE { get; set; }

        public static est_data_class FromCsv(string csvLine)
        {
            string[] values = csvLine.Split(',');
            est_data_class est_data = new est_data_class();

            est_data.Size = values[0];
            est_data.famName = values[1];
            est_data.ND1 = values[2];
            est_data.ND2 = values[3];
            est_data.ND3 = values[4];
            est_data.ND4 = values[5];
            est_data.KDS_MfrPart = values[6];
            est_data.KDS_HPH = values[7];
            est_data.KDS_MfrList = values[8];
            est_data.KDS_MCAA_LBR_RATE = values[9];
            est_data.KDS_LBR_RATE = values[10];

            return est_data;
        }  // End From Csv
        public override string ToString()
        {
            return "" +
                " Size: " + Size +
                " :: famName = " + famName +
                " :: ND1 = " + ND1 +
                " :: ND2 = " + ND2 +
                " :: ND3 = " + ND3 +
                " :: ND4 = " + ND4 +
                " :: KDS_MfrPart = " + KDS_MfrPart +
                " :: KDS_HPH = " + KDS_HPH +
                " :: KDS_MfrList = " + KDS_MfrList +
                " :: KDS_MCAA_LBR_RATE = " + KDS_MCAA_LBR_RATE +
                " :: KDS_LBR_RATE = " + KDS_LBR_RATE;
        }  // End of ToString

        public string ToFormatString()
        {
            return "" +
                " Size: " + Size +
                "\n famName = " + famName +
                "\n ND1 = " + ND1 +
                "\n ND2 = " + ND2 +
                "\n ND3 = " + ND3 +
                "\n ND4 = " + ND4 +
                "\n KDS_MfrPart = " + KDS_MfrPart +
                "\n KDS_HPH = " + KDS_HPH +
                "\n KDS_MfrList = " + KDS_MfrList +
                "\n KDS_MCAA_LBR_RATE = " + KDS_MCAA_LBR_RATE +
                "\n KDS_LBR_RATE = " + KDS_LBR_RATE;
        }  // End of ToFormatString
        public string ToCsvString()
        {
            return Size + "," +
                ND1 + "," + ND2 + "," + ND3 + "," + ND4 + "," +
                KDS_MfrPart + "," + KDS_HPH + "," + KDS_MfrList + "," + KDS_MCAA_LBR_RATE + "," + KDS_LBR_RATE;
        }  // End of ToCsvString

        public string ToCsvString_excld(string prop)
        {

            string props = this.ToCsvString();


            //return props.Replace(prop,"");
            return "," +
                ND1 + "," + ND2 + "," + ND3 + "," + ND4 + "," +
                KDS_MfrPart + "," + KDS_HPH + "," + KDS_MfrList + "," + KDS_MCAA_LBR_RATE + "," + KDS_LBR_RATE;
        }  // End of ToCsvString



        public string ToCsvString_wfname()
        {
            return Size + "," + famName + "," +
                ND1 + "," + ND2 + "," + ND3 + "," + ND4 + "," +
                KDS_MfrPart + "," + KDS_HPH + "," + KDS_MfrList + "," + KDS_MCAA_LBR_RATE + "," + KDS_LBR_RATE;
        }  // End of ToCsvString

    }  // End of Class est_data_class

    #endregion  //est_data_class


    #region  //  familyLoadOptions
    class familyLoadOptions : IFamilyLoadOptions
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

    #endregion  // End Of familyLoadOptions



}  // End of Namespace KDS_Module