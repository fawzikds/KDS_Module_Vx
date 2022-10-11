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
        const string csvFilePath_const = "Z:\\BIM\\Families\\SupplierCode\\CI_NH_fittings_Charlotte_KDS.csv";    // COntains HPH and List Price for all Elements used by KDS, fittings, Accessories, Finish, etc..
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
            #endregion // Initializations, Filters, Collections //

            #region   // Add the Category of Families to use (Pipe Fittings, Accessories, Fixtures...)
            List<BuiltInCategory> bic_lst = new List<BuiltInCategory>();

            bic_lst.Add(BuiltInCategory.OST_PipeFitting);
            bic_lst.Add(BuiltInCategory.OST_PlumbingFixtures);
            bic_lst.Add(BuiltInCategory.OST_PipeAccessory);
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
                bicFam.bic = bic.ToString();
                bicFam.fam_lst = famMain_lst;  //DstnctFam_lst;
                bicFam_lst.Add(bicFam);

                // Load List for Distinct family nmaes per bic
                bicFam_strct bicDstnctFam = new bicFam_strct();
                bicDstnctFam.bic = bic.ToString();
                bicDstnctFam.fam_lst = DstnctFam_lst;
                bicDstnctFam_lst.Add(bicDstnctFam);

                bicfam_tds += "\n BIC: " + bicFam.bic + ":: Count = " + bicFam.fam_lst.Count();
                //TaskDialog.Show("Execute", "BIC: " + bicFam.famTyp + ":: Count = " + bicFam.fam_lst.Count()+ "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");
            }
            //TaskDialog.Show("Execute", bicfam_tds + "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");

            #endregion // End Of Loop Thru every Category of Families and get all of its Elements in the actvDocument


            #region   // Loop thru every Category and Create the Dir Tree if it does not exist
            if (bicFam_lst.Count() > 0)
            {
                string ftf_tds = "famSmbl_lst.Count: " + bicFam_lst.Count();
                foreach (bicFam_strct bicfam in bicFam_lst)
                {
                    ftf_tds += "\n Family BIC: " + bicfam.bic + ":: Count: " + bicfam.fam_lst.Count;
                    if (bicfam.fam_lst.Count > 0)
                    {
                        CreateAndFill_FamilyDirTree(uidoc, bicfam.bic, bicfam.fam_lst);
                    }
                }  // End of for each FamTyFam ftf
                //TaskDialog.Show("Execute", ftf_tds + "\n\n\n\n\n\n\n\n\n\n\n\n");
            }// End if bicFam_lst is >0
            #endregion   // End OF Loop thru every Category and Create the Dir Tree if it does not exist


            #region   // Loop thru every Category and show Distinct Families
            if (bicDstnctFam_lst.Count() > 0)
            {
                string ftf_tds = "bicDstnctFam.Count: " + bicFam_lst.Count();
                foreach (bicFam_strct bicDstnctFam in bicDstnctFam_lst)
                {

                    ftf_tds = "\n Family BIC: " + bicDstnctFam.bic + ":: Count: " + bicDstnctFam.fam_lst.Count;
                    if (bicDstnctFam.fam_lst.Count > 0)
                    {
                        int ftf_bic_int = 0;
                        foreach (Family dstnctfam in bicDstnctFam.fam_lst)
                        {
                            ftf_bic_int++;
                            ftf_tds += "\n" + ftf_bic_int + "- " + dstnctfam.Name;
                        }
                        //TaskDialog.Show("Execute", ftf_tds + "\n\n\n\n\n\n\n\n\n\n\n\n");
                        Csv_to_DB(uidoc, csvFilePath_const, bicDstnctFam.bic, bicDstnctFam.fam_lst);   // Convert CVS to a Structured Data
                    }
                }  // End of for each FamTyFam ftf

            }// End if bicFam_lst is >0
            #endregion   // End OF Loop thru every Category and Create the Dir Tree if it does not exist

            /*
            #region // Split up CSV file //
            if (bicDstnctFam_lst.Count() > 0)
            {
                foreach (bicFam_strct bicDstnctFam in bicDstnctFam_lst)
                {
                    string ftf_tds = "\n Family BIC: " + bicDstnctFam.bic + ":: Count: " + bicDstnctFam.fam_lst.Count;
                    if (bicDstnctFam.fam_lst.Count > 0)
                    {

                        foreach (Family dstnctfam in bicDstnctFam.fam_lst)
                        {
                            TaskDialog.Show("Execute", "bic:" + bicDstnctFam.bic);
                            if (bicDstnctFam.fam_lst.Count > 0)
                            {
                                Csv_to_DB(uidoc, csvFilePath_const, bicDstnctFam.bic, bicDstnctFam.fam_lst);   // Convert CVS to a Structured Data
                            }
                        }
                    }
                }
            }
            #endregion
            */
            return Result.Succeeded;  //=====================================================


            #region // Import csv in families //
            foreach (Family fam in DstnctFam_lst)
            {
                ImportSizeLookUpTable(uidoc, fam);
                Create_sharedParameter_inFamily(uidoc, fam);
            }
            #endregion

            return Result.Succeeded;
        }



        #region   // Create and Fill Families in Dir tree

        public void CreateAndFill_FamilyDirTree(UIDocument uiDoc, String bic, List<Family> DstnctFam_lst)
        {

            Autodesk.Revit.DB.Document actvDoc = uiDoc.Document;

            foreach (Family fam in DstnctFam_lst)
            {
                string folderName = null;
                string input = fam.Name;

                //string famClass = fam.get_Parameter(BuiltInParameter.fam);
                //BuiltInParameter _bip = BuiltInParameter.OMNICLASS_CODE;


                int j = input.IndexOf("-");

                if (j >= 0)
                {
                    folderName = input.Substring(0, j);
                }

                //string famTypeDir = sharedParameter_dir + BuiltInCategory.OST_PipeFitting.ToString() + "\\";     // z:\BIM\KDS_TEMPLATE\2022\fittings
                string famTypeDir = sharedParameter_dir + bic + "\\";     // z:\BIM\KDS_TEMPLATE\2022\fittings
                string famManfDir = famTypeDir + folderName + "\\";                // z:\BIM\KDS_TEMPLATE\2022\fittings\KDS_Char_CI_NH
                string famDir = famManfDir + fam.Name + "\\";                      // z:\BIM\KDS_TEMPLATE\2022\fittings\KDS_Char_CI_NH\KDS_Char_CI_NH_Coupling\
                //string file = sharedParameter_dir + "fittings\\" + folderName + "\\" + fam.Name + "\\" + fam.Name + ".rfa";   // z:\BIM\KDS_TEMPLATE\2022\fittings\KDS_Char_CI_NH\KDS_Char_CI_NH_Coupling\KDS_Char_CI_NH_Coupling.rfa
                string file = famDir + fam.Name + ".rfa";
                /* DEBUG  TaskDialog.Show("CreateAndFill_FamilyDirTree", "filename = " + fam.Name +
                                            "\n famTypeDir = " + famTypeDir +
                                            "\n famManfDir = " + famManfDir +
                                            "\n famDir = " + famDir +
                                            "\n file = " + file + "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");     // The \n makes the dialog box longer, which in turns makes it wider so it can fit the file names.
 */
                //If directory does not exist, create it
                if (!Directory.Exists(famDir)) { Directory.CreateDirectory(famDir); }

                // if file Does not exists or if it is older than curren... recreate it
                if (!System.IO.File.Exists(file))//|| (System.IO.File.GetCreationTime(file) < DateTime.UtcNow.AddDays(-1)))
                {
                    try
                    {
                        Autodesk.Revit.DB.Document currSelFam = actvDoc.EditFamily(fam);

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
                }  // End of if File Exists or older

            }   // End For each Family
        }  // End of CreateAndFill_FamilyDirTree

        #endregion   // End Of Create and Fill Families in Dir tree



        #region  // Convert a CVS File to a Structured Data so we can work with.DstnctFam_lst
        // we have to use the Distinct Family list, (DstnctFam_lst) since we will have multiple lines that
        // represents different fitting sizes for each family that will need to go into the same csv file.
        public void Csv_to_DB(UIDocument uidoc, string csvFilePath_const, string famBIC_Nm, List<Family> DstnctFam_lst)
        {



            #region // test input lst
            string ftf_tds = "\n Family BIC: " + famBIC_Nm + ":: Count: " + DstnctFam_lst.Count;
            if (DstnctFam_lst.Count > 0)
            {
                int ftf_bic_int = 0;
                foreach (Family dstnctfam in DstnctFam_lst)
                {
                    ftf_bic_int++;
                    ftf_tds += "\n" + ftf_bic_int + "- " + dstnctfam.Name;
                }
                //TaskDialog.Show("Csv_to_DB", ftf_tds + "\n\n\n\n\n\n\n\n\n\n\n\n");
            }
            else { TaskDialog.Show("Csv_to_DB", "Distinct Family List Count is: " + DstnctFam_lst.Count); }

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
                //TaskDialog.Show("csv_DB", "BIC : " + famBIC_Nm + "\n  - Formatted Header: \n\n " + est_db_hdr.ToFormatString() + "\n\n - CSV Format: " + est_db_hdr.ToCsvString());

                est_data_class_lst = File.ReadAllLines(csvFilePath_const)
                                                        .Skip(1)
                                                        .Select(v => est_data_class.FromCsv(v))
                                                        .ToList();
                //TaskDialog.Show("csv_DB", "BIC : " + famBIC_Nm + "\n - Formatted Data:  \n\n" + est_data_class_lst[2].ToFormatString() + "\n\n - CSV Format: " + est_data_class_lst[2].ToCsvString());
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
                tds_edc += "\n " + edc_cnt + "- Formatted Data: " + est_data_class.ToCsvString();
            }
            TaskDialog.Show("Csv_to_DB", tds_edc);
            #endregion  // Get Data from the KDS Estimation CSV File.



            #region  // Loop Thru every Distinct Family and get its associated Rows from the Data var (est_db_data); Then Save these rows to a csv file.
            string sb1 = null;
            string dstnctFam_summary = " Distinct Family Summary for Category: " + famBIC_Nm + "\n";
            int dstnctFam_summary_cnt = 0;

            #region // Loop Thru every Distinct Family and get its associated Rows from the Data var (est_db_data);
            foreach (Family DstnctFam in DstnctFam_lst)
            {
                string DstnctFam_nm = DstnctFam.Name;
                //TaskDialog.Show("Csv_to_DB", "BIC: " + famBIC_Nm + ":: Currrent Distinct Family Name is: " + DstnctFam_nm + "\n\n\n\n\n\n\n\n\n\n");

                List<est_data_class> est_db_DstnctFam_lst = est_data_class_lst.Where(dc => dc.famName.Equals(DstnctFam_nm)).ToList<est_data_class>();    // Holds the ows per Family to save a csv



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
                        TDS += "\n" + est_db_fam.ToFormatString();
                    }

                    //TaskDialog.Show("Csv_to_DB", "TDS: " + TDS);

                    #endregion // End Of Loop Thru every Distinct Family and get its associated Rows from the Data var (est_db_data);

                    #region  // Save these rows to a csv file.region

                    string folderName = null;
                    string input = DstnctFam_nm;
                    int j = input.IndexOf("-");

                    if (j >= 0)
                    {
                        folderName = input.Substring(0, j);
                    }

                    string file = sharedParameter_dir + famBIC_Nm + "\\" + folderName + "\\" + DstnctFam.Name + "\\" + DstnctFam.Name + ".csv";
                    dstnctFam_summary += " -- CSV File Name:" + file;

                    //sb1 += file + "\n";
                    //string fileToCopy = "c:\\myFolder\\myFile.txt";
                    //string destinationDirectory = "c:\\myDestinationFolder\\";

                    //File.Copy(fileToCopy, destinationDirectory + Path.GetFileName(fileToCopy));




                    /*using (var stream = System.IO.File.CreateText(file))
                    {
                        //for (int i = 0; i < reader.Count(); i++)
                        foreach (var ds in est_db_fam)
                        {
                            //string first = reader[i].ToString();
                            //string second = image.ToString();
                            string csvRow = string.Format("{0},{1},{2},{3},{4},{5} ", ds.Size, ds.KDS_HPH, ds.KDS_MfrPart, ds.KDS_MCAA_LBR_RATE, ds.KDS_LBR_RATE, ds.KDS_MfrList);
                            //string.Join(",", line);
                            stream.WriteLine(csvRow);
                        }  // End Of foreach est_db_fam
                    }  // End Of Using


                    TaskDialog.Show("CSV_TO_DB", sb1);
                    */
                    #endregion  // End Of Save Found Famiy rows to a csv file.
                }  // end foreach famName

            }   TaskDialog.Show("Csv_to_DB", "TDS: " + dstnctFam_summary);
            #endregion  // Loop Thru every Distinct Family and get its associated Rows from the Data var (est_db_data); Then Save these rows to a csv file.


        }  // end of CSV_TO_DB
        #endregion  // End Of Convert a CVS File to a Structured Data so we can work with.DstnctFam_lst


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

        public void CreateAndSetFormula(ExternalDefinition eDef, UIDocument uidoc, Autodesk.Revit.DB.Document currSelFam)
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
            Autodesk.Revit.DB.Document actvDoc = uidoc.Document;
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

            Autodesk.Revit.DB.Document famDoc = fam.Document;

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

        #region // bicFam_strct
        public struct bicFam_strct
        {
            public string bic;
            public List<Family> fam_lst;

            public bicFam_strct(string bic, List<Family> fam_lst)
            {
                this.bic = bic;
                this.fam_lst = fam_lst;

            }

        } // End of bicFam_strct
        #endregion  // End of bicFam_strct


        #region // est_data_strct
        public struct est_data_strct
        {
            public string famName;
            public string Size;
            public string ND1;
            public string ND2;
            public string ND3;
            public string ND4;
            public string KDS_MfrPart;
            public string KDS_HPH;
            public string KDS_MfrList;
            public string KDS_MCAA_LBR_RATE;
            public string KDS_LBR_RATE;

            public est_data_strct(
                string famName,
            string Size,
            string ND1,
             string ND2,
             string ND3,
             string ND4,
             string KDS_MfrPart,
             string KDS_HPH,
             string KDS_MfrList,
             string KDS_MCAA_LBR_RATE,
             string KDS_LBR_RATE)
            {
                this.famName = famName;
                this.Size = Size;
                this.ND1 = ND1;
                this.ND2 = ND2;
                this.ND3 = ND3;
                this.ND4 = ND4;
                this.KDS_MfrPart = KDS_MfrPart;
                this.KDS_HPH = KDS_HPH;
                this.KDS_MfrList = KDS_MfrList;
                this.KDS_MCAA_LBR_RATE = KDS_MCAA_LBR_RATE;
                this.KDS_LBR_RATE = KDS_LBR_RATE;
            }

        } // End of est_data_strct
        #endregion  // End of est_data_strct






    }  //end of Class create_update_Fam_KDS_CSV
    #endregion  // create_update_Fam_KDS_CSV

    #region  //est_data_class
    class est_data_class
    {
        public string famName { get; set; }
        public string Size { get; set; }
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
                " famName: " + famName +
                " :: Size = " + Size +
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
                " famName: " + famName +
                "\n Size = " + Size +
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
            return famName + "," + Size + "," +
                ND1 + "," + ND2 + "," + ND3 + "," + ND4 + "," +
                KDS_MfrPart + "," + KDS_HPH + "," + KDS_MfrList + "," + KDS_MCAA_LBR_RATE + "," + KDS_LBR_RATE;
        }  // End of ToCsvString

    }  // End of Class est_data_class

    #endregion  //est_data_class



}  // End of Namespace KDS_Module