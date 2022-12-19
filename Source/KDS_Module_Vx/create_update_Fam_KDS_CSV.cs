using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Utility;

namespace KDS_Module_Vx
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
        //const string csvFilePath_const = "Z:\BIM\KDS_SUPPLIER_CODE\\KDS_CI_NH_BLK_MLBL.csv";    // COntains HPH and List Price for all Elements used by KDS, fittings, Accessories, Finish, etc..
        //const string csvFilePath_const = get_PriceDB_fileName();

#if (RVT2022)
        const string sharedParameter_dir = "Z:\\BIM\\KDS_TEMPLATE\\2022\\";
        const string sharedParameter_fn = sharedParameter_dir + "KDS_SHARED_PARAMS.txt";          // Shared Parmeters.. this is how we get into the family parameters. Still not sure if this is best to use.
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
            Autodesk.Revit.DB.Document actvDoc = uidoc.Document;

            List<FamilySymbol> famSmbl_lst = new List<FamilySymbol>();
            List<Family> famMain_lst = new List<Family>();
            List<Family> dstnctFam_lst = new List<Family>();
            List<bicFam_strct> bicFam_lst = new List<bicFam_strct>();   // holds list of all families per bic
            List<bicFam_strct> bicDstnctFam_lst = new List<bicFam_strct>();    // Holds lists of distinct famlies per bic
            string csvFilePath_const = get_PriceDB_fileName();
            #endregion // Initializations, Filters, Collections //

            #region   // Add the Category of Families to use (Pipe Fittings, Accessories, Fixtures...)
            List<BuiltInCategory> bic_lst = new List<BuiltInCategory>();

            bic_lst.Add(BuiltInCategory.OST_PipeFitting);
            bic_lst.Add(BuiltInCategory.OST_PlumbingFixtures);
            bic_lst.Add(BuiltInCategory.OST_PipeAccessory);
            #endregion  // End Of Add the Category of Families to use (Pipe Fittings, Accessories, Fixtures...)

            #region // Loop Thru every Category of Families(fitt,Fixt,Acces) and get all of its Elements in the actvDocument
            // Then store in bicFam structure so we can access all families by their Category.
            string bicfam_tds = " Family Count per BIC\n";
            foreach (BuiltInCategory bic in bic_lst)
            {
                famSmbl_lst = new FilteredElementCollector(uidoc.Document).OfCategory(bic).WhereElementIsElementType().Select(df => df as FamilySymbol).ToList();
                famMain_lst = famSmbl_lst.Select(fs => fs.Family).ToList();  // Get Family From familySymbols

                // This get a unique list of all families per Category.
                dstnctFam_lst = famMain_lst.GroupBy(f => f.Name).Select(g => g.First()).ToList();    //how to find unique families based on family name, not family type, of given categories

                // Load List for families per bic
                bicFam_strct bicFam = new bicFam_strct();
                bicFam.bic_str = bic.ToString();
                bicFam.fam_lst = famMain_lst;  //DstnctFam_lst;
                bicFam_lst.Add(bicFam);

                // Load List for Distinct family nmaes per bic
                bicFam_strct bicDstnctFam = new bicFam_strct();
                bicDstnctFam.bic_str = bic.ToString();
                bicDstnctFam.fam_lst = dstnctFam_lst;
                bicDstnctFam_lst.Add(bicDstnctFam);

                bicfam_tds += "\n BIC: " + bicFam.bic_str + ":: Count = " + bicFam.fam_lst.Count();
                //TaskDialog.Show("Execute", "BIC: " + bicFam.famTyp + ":: Count = " + bicFam.fam_lst.Count()+ "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");
            }
            //TaskDialog.Show("Execute", bicfam_tds + "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");

            #endregion // End Of Loop Thru every Category of Families and get all of its Elements in the actvDocument

            #region   // Loop thru every Category and Create the Dir Tree if it does not exist  including .rfa and its associated but empty .csv
            if (bicFam_lst.Count() > 0)
            {
                string ftf_tds = "bicFam_lst.Count: " + bicFam_lst.Count();
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
            //show_DstnctFam(bicDstnctFam_lst); // Loop thru every Category and show Distinct Families  -- Debug Only

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
                    imprt_tds = "bicDstnctFam.bic_str:  " + bicDstnctFam.bic_str + "\n\n";
                    //TaskDialog.Show("Execute", " Getting Ready to Import csvFiles for: " + imprt_tds);
                    if (bicDstnctFam.fam_lst.Count > 0)
                    {
                        imprt_tds += "bicDstnctFam.fam_lst.Count: " + bicDstnctFam.fam_lst.Count;
                        int ftf_cnt = 0;

                        foreach (Family dstnctFam in bicDstnctFam.fam_lst)
                        {
                            ftf_cnt++;
                            // Get RFA and CSV File Name
                            string dstnctFamPath = get_FileNames(sharedParameter_dir, bicDstnctFam.bic_str, dstnctFam.Name, ".rfa");
                            string csvFilePath = get_FileNames(sharedParameter_dir, bicDstnctFam.bic_str, dstnctFam.Name, ".csv");
                            string csvFileName = get_FileNames(sharedParameter_dir, bicDstnctFam.bic_str, dstnctFam.Name, "");
                            //TaskDialog.Show("ImportSizeLookUpTable", "\n Document File Name : " + dstnctFamPath + "\n With csvFilePath: " + csvFilePath);
                            if (dstnctFam.IsEditable)
                            {
                                Autodesk.Revit.DB.Document dstnctFamEdt = actvDoc.EditFamily(dstnctFam);

                                // Get FSTM (Family Size Table Manager)
                                imprt_tds += "\n " + ftf_cnt + "- " + dstnctFam.Name + ":: Errors: " + ImportSizeLookUpTable(bicDstnctFam.bic_str, dstnctFamEdt, dstnctFamPath, csvFilePath);

                                // This is a MAJOR function, i left it within this loop and not sepreate it out, since it is easier to load and save family right after i changed Things.
                                Create_sharedParameter_inFamily(actvDoc, bicDstnctFam.bic_str, dstnctFamEdt, sharedParameter_fn, dstnctFamPath);

                                dstnctFamEdt.LoadFamily(actvDoc, new familyLoadOptions()); // here the update begins
                                //TaskDialog.Show("ImportSizeLookUpTable", "after  LoadFamily ");

                                // Save family
                                SaveAsOptions saveoptions = new SaveAsOptions();
                                saveoptions.OverwriteExistingFile = true;

                                try
                                {
                                    dstnctFamEdt.SaveAs(dstnctFamPath, saveoptions);  // Error is here
                                                                                      //TaskDialog.Show("ImportSizeLookUpTable", "after  SaveAs ");
                                }
                                catch (Exception ex)
                                {
                                    TaskDialog.Show("ImportSizeLookUpTable", "Error SaveAs after Loading family. \n Exception: " + ex);
                                }

                                try
                                {
                                    dstnctFamEdt.Close(false);
                                    //TaskDialog.Show("ImportSizeLookUpTable", "after  Close ");
                                }
                                catch (Exception ex)
                                {
                                    TaskDialog.Show("ImportSizeLookUpTable", "Error Close after Loading family. \n Exception: " + ex);
                                }
                            } // End of if dstnctFam isEditable
                            else
                            {
                                TaskDialog.Show("create_sharedPamaeter_in_Family", "Family " + dstnctFam + "  is NOT editable");
                            }  // uneditable families

                        }  // foreach bicDstnctFam
                        TaskDialog.Show("Execute", imprt_tds + "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");
                    }  // End Of if  bicDstnctFam.fam_lst.Count > 0
                }  // End Of foreach bicDstnctFam in bicDstnctFam_lst
            }  // End Of if bicDstnctFam_lst.Count() > 0

            #endregion

            return Result.Succeeded;
        }


        #region   // Create and Fill Families in Dir tree  with the .rfa and the .csv (csv will be empty now)
        public void CreateAndFill_FamilyDirTree(UIDocument uiDoc, string famBIC_Nm, string csvFilePath_const, List<Family> dstnctFam_lst)
        {
            Autodesk.Revit.DB.Document actvDoc = uiDoc.Document;

            foreach (Family dstnctFam in dstnctFam_lst)
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
                string famDir = famManfDir + dstnctFam.Name + "\\";                      // z:\BIM\KDS_TEMPLATE\2022\fittings\KDS_Char_CI_NH\KDS_Char_CI_NH-Coupling\
                                                                                         //string file = sharedParameter_dir + "fittings\\" + folderName + "\\" + fam.Name + "\\" + fam.Name + ".rfa";   // z:\BIM\KDS_TEMPLATE\2022\fittings\KDS_Char_CI-NH\KDS_Char_CI_NH-Coupling\KDS_Char_CI_NH-Coupling.rfa

                string dstnctFamFname = get_FileNames(sharedParameter_dir, famBIC_Nm, dstnctFam.Name, ".rfa");


                /* DEBUG  TaskDialog.Show("CreateAndFill_FamilyDirTree", "filename = " + fam.Name +
                                            "\n famTypeDir = " + famTypeDir +
                                            "\n famManfDir = " + famManfDir +
                                            "\n famDir = " + famDir +
                                            "\n file = " + file + "\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n");     // The \n makes the dialog box longer, which in turns makes it wider so it can fit the file names.
 */
                //If directory does not exist, create it
                if (!Directory.Exists(famDir)) { Directory.CreateDirectory(famDir); }

                // if file Does not exists or if it is older than current... recreate it... this is now Always Overwrite 
                if (!System.IO.File.Exists(dstnctFamFname) || (System.IO.File.GetCreationTime(dstnctFamFname) < DateTime.UtcNow.AddMonths(-1)))
                {
                    try
                    {
                        Autodesk.Revit.DB.Document currSelFam = actvDoc.EditFamily(dstnctFam);

                        SaveAsOptions saveoptions = new SaveAsOptions();
                        saveoptions.OverwriteExistingFile = true;

                        using (Transaction saveFam_trx = new Transaction(actvDoc, "Save Family"))
                        {
                            saveFam_trx.Start();
                            currSelFam.SaveAs(dstnctFamFname, saveoptions);
                            currSelFam.Close(false);
                            saveFam_trx.Commit();
                            saveFam_trx.Dispose();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        TaskDialog.Show("CreateAndFill_FamilyDirTree", dstnctFamFname + " Exception: \n" + ex);
                    }
                }  // End of if File Exists or older
                #region  // Create an Empty csv for each family lookup table file.  --- Uneccessary
                //i don't need this since i will create it later as i am filling them with content from the KDS  est csv file.
                // However, since i am still in testing, i need to have a file for every family so i can test the importLookuptable function.
                // But this will not work for testing adding other parameters since the files would not have any content.... unless if i add the header only ???
                //Add CSV Files as Well
                //if file Does not exists or if it is older than curren... recreate it

                est_data_class est_db_hdr = est_data_class.FromCsv(System.IO.File.ReadAllLines(csvFilePath_const).First());
                string csvRow = "";
                // foreach (Family dstnctFam in dstnctFam_lst)
                // {
                string csvFile = get_FileNames(sharedParameter_dir, famBIC_Nm, dstnctFam.Name, ".csv");
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(csvFile));
                using (var stream = System.IO.File.CreateText(csvFile))     // Opens or Creates File if it does not exist. if It exists it will overwrite its content
                {
                    // Write Headers first
                    csvRow = est_db_hdr.ToCsvString_excld("Size");
                    stream.WriteLine(csvRow);
                } // End Of Using 
                  // }  // End Of foreach est_db_fam
                #endregion  // END Of Create an Empty csv for each family lookup table file.  --- Uneccessary
            }   // End For each Family
        }  // End of CreateAndFill_FamilyDirTree
        #endregion   // End Of Create and Fill Families in Dir tree


        #region  // Convert a CVS File to a Structured Data so we can work with.DstnctFam_lst
        // we have to use the Distinct Family list, (DstnctFam_lst) since we will have multiple lines that
        // represents different fitting sizes for each family that will need to go into the same csv file.
        public void Csv_to_DB(UIDocument uidoc, string csvFilePath_const, string famBIC_Nm, List<Family> dstnctFam_lst)
        {
            #region // test input lst   -- Debug Only  -- Commented Out
            /*          string ftf_tds = "\n Family BIC: " + famBIC_Nm + ":: Count: " + dstnctFam_lst.Count;
                      if (DstnctFam_lst.Count > 0)
                      {
                          int ftf_bic_int = 0;
                          foreach (Family dstnctfam in dstnctFam_lst)
                          {
                              ftf_bic_int++;
                              ftf_tds += "\n" + ftf_bic_int + "- " + dstnctfam.Name;
                          }
                          TaskDialog.Show("Csv_to_DB", ftf_tds + "\n\n\n\n\n\n\n\n\n\n\n\n");
                      }
                      else { TaskDialog.Show("Csv_to_DB", "Distinct Family List Count is: " + dstnctFam_lst.Count); }
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
            foreach (Family dstnctFam in dstnctFam_lst)
            {
                string dstnctFam_nm = dstnctFam.Name;
                //TaskDialog.Show("Csv_to_DB", "BIC: " + famBIC_Nm + ":: Currrent Distinct Family Name is: " + dstnctFam_nm + "\n\n\n\n\n\n\n\n\n\n");

                List<est_data_class> est_db_DstnctFam_lst = est_data_class_lst.Where(dc => dc.famName.Equals(dstnctFam_nm)).ToList<est_data_class>();    // Holds the rows per Family to save a csv

                if (est_db_DstnctFam_lst.Count > 0)
                {
                    dstnctFam_summary_cnt++;
                    dstnctFam_summary += "\n " + dstnctFam_summary_cnt + "- Currrent Distinct Family " + dstnctFam_nm + " count: " + est_db_DstnctFam_lst.Count;

                    //dstnctFam_summary += "\n Currrent Distinct Family " + dstnctFam_nm + " count: " + est_db_DstnctFam_lst.Count;
                    //TaskDialog.Show("Csv_to_DB", "BIC: " + famBIC_Nm + ":: Currrent Distinct Family " + dstnctFam_nm + " count: " + est_db_DstnctFam_lst.Count + "\n\n\n\n\n\n\n\n\n\n");
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

                    string dstnctFamFname = get_FileNames(sharedParameter_dir, famBIC_Nm, dstnctFam.Name, ".csv");

                    dstnctFam_summary += " -- CSV File Name:" + dstnctFamFname;

                    #region  // Save these rows to a csv file // Opens or Creates File if it does not exist. if It exists it will overwrite its content
                    using (var stream = System.IO.File.CreateText(dstnctFamFname))     // Opens or Creates File if it does not exist. if It exists it will overwrite its content
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
        public bool Create_sharedParameter_inFamily(Document actvDoc, string famBIC_Nm, Document dstnctFamEdt, string sharedParameter_fn, string dstnctFamPath)
        {
            //TaskDialog.Show("test", "current family: " + family.Name);
            /// Hard Coded Formulas ///   I need this to sort other lists by the order seen here... so not alphabetical, but by another list.
            List<string> paramFixedName_lst = new List<string> { "KDS_ID_tbl", "KDS_ND0", "KDS_ND1", "KDS_ND2", "KDS_ND3", "KDS_HPH", "KDS_MfrList", "KDS_MfrPart", "KDS_MCAA_LBR_RATE", "KDS_LBR_RATE" };

            /// Hard Coded Formulas: Set a generic values for all parameters except for the lookuptable and for the NDs- This i will get from associated Parameter name to a ConnectorElement.
            List<famParam_Data_class> fpData_lst = new List<famParam_Data_class>();
            fpData_lst = Get_HardCodedFamParamFormulas();

            #region // List all  names of fpData_lst returned from Get_HardCodedFamParamFormulas-- Confirmed - Debug Only -- Commented Out
            /*string fpData_Name_str = "list of all fpData_Names\n";

            foreach (famParam_Data_class fpD in fpData_lst)
            {
                fpData_Name_str += "\n - fpData Name: " + fpD.Name;
            }
            TaskDialog.Show("Create_sharedParameter_inFamily", fpData_Name_str);*/
            #endregion   // End Of Debug - List all  names of fpData_lst returned from Get_HardCodedFamParamFormulas-- Confirmed

            #region // Get a customSorted list of All External Definition in Shared Parameter File. KDS_ID_tble and KDS_NDs params should 
            //be processed before hph, mfrPart, mfrList, and MCAA_LBR and LBR rates.

            // Data Class to hold the different pieces of a parameter, Name, type, default Value and Formula
            famParam_Data_class fpData = new famParam_Data_class();

            // Set Filename for SharedParameters and open file
            actvDoc.Application.SharedParametersFilename = sharedParameter_fn;
            DefinitionFile defFile = actvDoc.Application.OpenSharedParameterFile();

            // Get List of Groups in SharedParameter File.
            DefinitionGroups myGroups = defFile.Groups;
            IList<DefinitionGroup> myGroup_lst = myGroups.ToList();

            // Get Group Definitions.  Definitions an be: 
            //   InternalDefinitions that reside inside a Revit Document, OR
            //   ExternalDefinintions that reside in an external File...such as shared Parameters. we want these.
            Autodesk.Revit.DB.Definitions myDefinitions = myGroup_lst[0].Definitions;
            List<Definition> Def_lst = myDefinitions.ToList();   //    get_Item("PRL_WIDTH") as ExternalDefinition;

            //  Cast Definitions to ExternalDefinitions..... How did myGroup become a shared parameter. 
            var extDef_lst = Def_lst.Select(df => df as ExternalDefinition).ToList();

            // Sort The list of Shared Parameter by the my paramFixedName_lst since we oder is important in when i create the parameters.
            List<ExternalDefinition> extDef_sorted_lst = sort_listByAnother(extDef_lst, paramFixedName_lst);
            #endregion // End Of Get a customSorted list of All External Definition in Shared Parametr File


            // Get  Family Manager 
            FamilyManager dstnctFamMngr = dstnctFamEdt.FamilyManager;

            // Get List of all ConnectoElements in family.
            List<Element> connElmnt_lst = new FilteredElementCollector(dstnctFamEdt).OfCategory(BuiltInCategory.OST_ConnectorElem).ToElements().ToList();

            // Get List of all Parameters in family.
            List<FamilyParameter> dstnctFamParam_lst = dstnctFamMngr.GetParameters().ToList<FamilyParameter>();

            // Find All external Definition Shared Parameters that are already in the list of Family parameters
            List<FamilyParameter> dstnctFamExtDefParam_lst = dstnctFamParam_lst.Where(a => extDef_lst.Any(b => b.Name == a.Definition.Name)).ToList();

            // Find all External Definition Shared Parameters that are missing from the List of Family Parmaters
            List<ExternalDefinition> missingFamParams_lst = extDef_lst.Where(a => !dstnctFamExtDefParam_lst.Any(b => b.Definition.Name == a.Name)).ToList();

            // Map ND parameter Formulas to the connectorElments' Associated Parameter Name.


            #region   Loop through list of exDef_lst parameters ( well i will use the new tuple so i can control the order of creating parameters, which is important)
            // i can either do :
            //    - Load fpData From the KDS_EST_DB file so we can use it in 2 loops below to create or fill parameters.
            //    - We may need to pass it as an argument here instead of calling for it by some function.
            // OR i can:
            //    - Loop through list of exDef_lst parameters,  and try to create it. if success, if not still good. ( i don't think there is an internal isExist)
            //    - Set the values/ formulas of each.... 
            //    - This is better, since i have to create things per some order... KDS_ID_tbl, then the other parameters, prefereably ND1 thru ND4 are in that order.


            #region // LOOP Thru all extDef, and Delete old nonshared parameter with same name, and to create new shared parameters as from extDef.
            string foundfamParam_str = " List Of All Found Parameters For: " + dstnctFamEdt.Title + "\n";
            int foundfamParam_int = 0;
            string missingfamParam_str = " List Of All Missing Parameters For: " + dstnctFamEdt.Title + "\n";
            int missingfamParam_int = 0;

            foreach (ExternalDefinition extDef in extDef_sorted_lst)
            {
                if (extDef.Name != null)
                {
                    FamilyParameter tmpParam = dstnctFamMngr.get_Parameter(extDef.Name);

                    // If it Exists but not a "Shared Parameter"  recreate as shared. since i am using Shared parameters and older non-shared params need to go.
                    if (tmpParam != null)
                    {
                        Transaction del_param_trx = new Transaction(dstnctFamEdt, "Delete Parameter");
                        del_param_trx.Start();
                        try
                        {
                            dstnctFamMngr.RemoveParameter(tmpParam);
                        }
                        catch (System.Exception ex)
                        {
                            TaskDialog.Show("Create_sharedParameter_inFamily", "Exception \n- Could Not Delete Parameter Name: " + extDef.Name + " \n- Family Name: " + dstnctFamEdt.Title + "\n Exception: " + ex);
                        }

                        del_param_trx.Commit();
                        del_param_trx.Dispose();

                        Transaction cr_param_trx = new Transaction(dstnctFamEdt, "Create Parameter");
                        cr_param_trx.Start();
                        try
                        {
                            tmpParam = dstnctFamMngr.AddParameter(extDef, BuiltInParameterGroup.PG_IDENTITY_DATA, true);
                            //TaskDialog.Show("Create_sharedParameter_inFamily", "Re-Created : " + tmpParam.Definition.Name);
                        }
                        catch (System.Exception ex)
                        {
                            TaskDialog.Show("Create_sharedParameter_inFamily", "Exception \n- Create Parameter Name: " + extDef.Name + " \n- Family Name: " + dstnctFamEdt.Title + "\n Exception: " + ex);
                        }

                        cr_param_trx.Commit();
                        cr_param_trx.Dispose();

                        foundfamParam_int++;
                        foundfamParam_str += foundfamParam_int + "- Param Name: " + extDef.Name + "\n";
                    }
                    // If parameter DOES NOT  Exists  ... Create it...
                    if (tmpParam == null)
                    {
                        Transaction cr1_param_trx = new Transaction(dstnctFamEdt, "Create Values and Formulas");
                        cr1_param_trx.Start();
                        try
                        {
                            tmpParam = dstnctFamMngr.AddParameter(extDef, BuiltInParameterGroup.PG_IDENTITY_DATA, true);
                            //TaskDialog.Show("Create_sharedParameter_inFamily", "Re-Created : " + tmpParam.Definition.Name);
                        }
                        catch (System.Exception ex)
                        {
                            TaskDialog.Show("Create_sharedParameter_inFamily", "Exception \n- Create Parameter Name: " + extDef.Name + " \n- Family Name: " + dstnctFamEdt.Title + "\n Exception: " + ex);
                        }

                        cr1_param_trx.Commit();
                        cr1_param_trx.Dispose();
                        missingfamParam_int++;
                        missingfamParam_str += missingfamParam_int + "- Missing Param Name: " + extDef.Name + "\n";
                        //TaskDialog.Show("Create_sharedParameter_inFamily", "created : " + tmpParam.Definition.Name);
                    }
                }  // if extDef is not null
            }  // foreach dstnctFamExtDefParam
               //TaskDialog.Show("Create_sharedParameter_inFamily", foundfamParam_str);
               //TaskDialog.Show("missingFamParams", missingfamParam_str);
            #endregion // End Of LOOP Thru all extDef, and Delete old nonshared parameter with same name, and to create new shared parameters as from extDef.


            #region// LOOP thru extDef parameters and add parmeter formulas 

            //string assignedfamParam_str    = " List Of All Found Parameters For: " + dstnctFamEdt.Title + "\n";   // for Debug
            //string notAssignedfamParam_str = " List Of All Missing Parameters For: " + dstnctFamEdt.Title + "\n"; //for Debug
            // This String will expand to include any Parmeter Names that are associated to a ConnectorElement. This way include only these in the Formula string.
            string ND_postFix_str = "";    // ",KDS_ND0, KDS_ND1, KDS_ND2, KDS_ND3)\"";

            foreach (ExternalDefinition extDef in extDef_sorted_lst)
            {

                if (extDef.Name != null)
                {
                    FamilyParameter tmpParam = dstnctFamMngr.get_Parameter(extDef.Name);

                    // If it Exists but not a "Shared Parameter"  recreate as shared. since i am using Shared parameters and all older non-shared params need to go.
                    if (tmpParam != null && tmpParam.CanAssignFormula)
                    {
                        string formula_str = "";
                        string vlu_str = "";

                        // Set the KDS_NDs Parameter Formula Value. This is based on the associated Parameters to Connector Elements. see inner loop
                        if (extDef.Name.Contains("KDS_ND"))
                        {
                            // If extDef is true and extDef.Name is true and extDef.Length !=0 is also true,
                            // then return the last character of extDef.Name ohterwise return 0... but convert to string
                            string ce_index_str = (extDef?.Name?.Length != 0 ? extDef.Name[extDef.Name.Length - 1] : '0').ToString();

                            // Convert the string to integer... Possible values 0, 1, 2, 3
                            int ce_index = Int32.Parse(ce_index_str);

                            // Possible Counts of ConnectorElements are: 1, 2, 3, 4 (for a Corss e.g.)
                            if (ce_index < connElmnt_lst.Count)
                            {
                                //TaskDialog.Show("Create_sharedParameter_inFamily", "Assigning Formula:" + "\n- Family Name: " + dstnctFamEdt.Title + "\n- Formula Value: " + formula_str + "     -||");
                                ND_postFix_str += ", " + extDef.Name;
                                // Add actual Formula
                                formula_str = get_AssociatedParametersofConnectorElement(dstnctFamEdt, connElmnt_lst[ce_index] as ConnectorElement);
                                add_FamParam_Formula(dstnctFamEdt, dstnctFamMngr, tmpParam, formula_str + "* 2");
                            }
                            else  // If this tmpParameter does not have an Associated ConnectorElement, then Delete that Parmeter from the Family.
                                  // It is nicer to delete it out of the lookuptable.csv as well, but it is not necessary.
                            {
                                Transaction remove_param_trx = new Transaction(dstnctFamEdt, "Add Parameter Values");
                                remove_param_trx.Start();
                                try
                                {
                                    dstnctFamMngr.RemoveParameter(tmpParam);
                                    //add_FamParam_Formula(dstnctFamEdt, dstnctFamMngr, tmpParam, null);  // Erase Current Formula 
                                }
                                catch (Exception ex)
                                {
                                    TaskDialog.Show("Create_sharedParameter_inFamily", "Remove tmpParam:" + "\n- Family Param Name: " + tmpParam.Definition.Name + "\n- Exception: " + ex);
                                }
                                remove_param_trx.Commit();
                                remove_param_trx.Dispose();
                            }
                        }

                        else if (extDef.Name == "KDS_ID_tbl")
                        {
                            // !!!!Family Types Special Handling.!!!!
                            // e.g. we have Hammer Arrestor with inlet Sizes of 1/2, 3/4, 1/ 2... each is not an instance, but a Type.
                            // If there are more than one Type in a Family, then i have to loop thru evrey Family Type and assign the size_Lookup Table (KDS_ID_tbl) value.
                            // Surprisingly, I did not need to do that for every formula though, but only for the lookuptable parameters. (KDS_ID_tbl)
                            // Adding a formula for one Family Type Filled it for all familytypes, which not intuitive. You'd think i need to do it for all Parameters.
                            FamilyTypeSet famTypes_set = dstnctFamMngr.Types;

                            foreach (FamilyType famType in famTypes_set)
                            {
                                dstnctFamMngr = SetFamilyManager_CurrentFamilyType(dstnctFamEdt, dstnctFamMngr, famType);

                                Transaction set_Lookup_param_trx = new Transaction(dstnctFamEdt, "Add Parameter Values");
                                set_Lookup_param_trx.Start();
                                try
                                {
                                    vlu_str = get_FileNames(sharedParameter_dir, famBIC_Nm, dstnctFamEdt.Title, "");
                                    //TaskDialog.Show("Create_sharedParameter_inFamily", "Assigning Formula:" + "\n- Family Name: " + dstnctFamEdt.Title + "\n- Value: " + vlu_str + "     -||");
                                    //dstnctFamMngr.Set(tmpParam, vlu_str);
                                    dstnctFamMngr.Set(tmpParam, dstnctFamEdt.Title);   // we only need the filename without its Path nor its Extension

                                    set_Lookup_param_trx.Commit();
                                    set_Lookup_param_trx.Dispose();
                                }   // Loop Thru Family Types 

                                catch (System.Exception ex)
                                {
                                    /*TaskDialog.Show("Create_sharedParameter_inFamily", "Exception \n- Could Not Get Associated Parameter Name For extDef: " + extDef.Name +
                                        " \n- Family Name: " + dstnctFamEdt.Title +
                                        " \n- Formula: " + formula_str +
                                        " \n- Exception: " + ex);*/
                                }
                            }

                        }  // End Of  If KDS_ID_tbl
                        else   // All Other Parameters other than the NDs and the KDS_ID_tbl
                        {
                            // Get the Root Part of the Formula String
                            formula_str = get_famParam_Data_Str(fpData_lst, extDef.Name);

                            // Add the ND index parametters to the formula.
                            formula_str += ND_postFix_str + ")";
                            //TaskDialog.Show("Create_sharedParameter_inFamily", "Assigning Formula:" + "\n- Family Name: " + dstnctFamEdt.Title + "\n- Formula Value: " + formula_str + "     -||");
                            //add_FamParam_Formula(dstnctFamEdt, dstnctFamMngr, tmpParam, null);
                            add_FamParam_Formula(dstnctFamEdt, dstnctFamMngr, tmpParam, formula_str);
                        }

                        foundfamParam_int++;
                        foundfamParam_str += foundfamParam_int + "- Param Name: " + extDef.Name + "\n";
                    }// if fampar.CanAssignFormula
                    else
                    {
                        //TaskDialog.Show("Create_sharedParameter_inFamily", "Exception \n- Create Parameter Name: " + extDef.Name + " \n- Family Name: " + dstnctFamEdt.Title);

                        missingfamParam_int++;
                        missingfamParam_str += missingfamParam_int + "- Missing Param Name: " + extDef.Name + "\n";
                        //TaskDialog.Show("Create_sharedParameter_inFamily", "created : " + tmpParam.Definition.Name);
                    }
                }  // if extDef is not null
            }  // foreach dstnctFamExtDefParam
            //TaskDialog.Show("Create_sharedParameter_inFamily", assignedfamParam_str);
            //TaskDialog.Show("missingFamParams", notAssignedfamParam_str);
            #endregion// End Of LOOP thru extDef parameters and add parmeter formulas 


            #endregion // End Of Loop through List of  of Found Family Parameters and set them to new Values


            return true;
        }  // End of create_sharedPamaeter_in_Family
        #endregion  // End Of Create_sharedParameter_inFamily 


        #region // Set FamilyManger Curent Type
        public FamilyManager SetFamilyManager_CurrentFamilyType(Document dstnctFamEdt, FamilyManager dstnctFamMngr, FamilyType famType)
        {
            using (Transaction trx = new Transaction(dstnctFamEdt))
            {
                trx.Start("Create Formulas");
                try
                {
                    dstnctFamMngr.CurrentType = famType;
                }
                catch (System.Exception ex)
                {
                    /*TaskDialog.Show("Create_sharedParameter_inFamily", "Exception \n- Could Not Set famManger.CurrentType=familyType : " +
                        " \n- Family Name: " + dstnctFamEdt.Title +
                        " \n- Family Type Name: " + famType.Name +
                        " \n- Exception: " + ex);*/
                }
                trx.Commit();
                trx.Dispose();
            }
            return dstnctFamMngr;

        }  // End Of SetFamilyManager_CurrentFamilyType
        #endregion // End Of Set FamilyManger Curent Type

        #region  // This seems not to be true in Revit 2022......I am leaving it here, just in case i needed it for other versions.
                 // Convoluted Way to set a formula ... similar to my familysizetable, you have to touch it first, before you can add a value to it.
        public void add_FamParam_Formula(Document dstnctFamEdt, FamilyManager dstnctFamMngr, FamilyParameter famParam, string formula_str)
        {

            using (Transaction trx = new Transaction(dstnctFamEdt))
            {
                trx.Start("Create Formulas");
                try
                {
                    if (famParam.CanAssignFormula)
                    {
                        //if (formula == "")
                        //   dstnctFamMngr.SetFormula(item, "1");
                        //else 
                        dstnctFamMngr.SetFormula(famParam, formula_str);

                        //if (item.Formula != null && item.Formula.ToString().StartsWith("1"))
                        // dstnctFamMngr.SetFormula(item, null);
                    }
                }
                catch (System.Exception ex)
                {
                    /* TaskDialog.Show("Create_sharedParameter_inFamily", "Exception \n- Could Not Get Associated Parameter Name For extDef: " + famParam.Definition.Name +
                         " \n- Family Name: " + dstnctFamEdt.Title +
                         " \n- Formula: " + formula_str +
                         " \n- Exception: " + ex);
                 */
                }
                trx.Commit();
                trx.Dispose();
            }

        }  // End Of  AddFormula
        #endregion  //End Of Convoluted Way to set a formula ... similar to my familysizetable, you have to touch it first, before you can add a value to it.

        #region // ImportSizeLookUpTable to Families
        // Imports the CSV lookup table file into a family.
        // It TURNS OUT that ......The editmanager of a family does not gurantee the existance of a familysizetable manager.
        //That is why you have to check for it first, then create it if it is null.
        // After you create the familysizetable manager, then you can use it to handle the lookup csv files you want ot import, export or delete.
        public string ImportSizeLookUpTable(string bic_nm, Document dstnctFamEdt, string dstnctFamPath, string csvFilePathName)
        {
            string lookupFileName = dstnctFamEdt.Title;  // csvFilePathName.Split('\\').Last();
            try
            {
                string importError = "";
                // Get the Family's Size and Table Manager

                Autodesk.Revit.DB.FamilySizeTableManager fstm = Autodesk.Revit.DB.FamilySizeTableManager.GetFamilySizeTableManager(dstnctFamEdt, dstnctFamEdt.OwnerFamily.Id);

                // If FSTM is Null, then it does not exist, So create it.
                if (fstm == null)
                {
                    //TaskDialog.Show("ImportSizeLookUpTable", "Starting with: " + "\n BIC: " + bic_nm + "\n dstnctFam.Name: " + dstnctFam.Name + "\n Creating an FSTM");
                    bool fstmResult = Autodesk.Revit.DB.FamilySizeTableManager.CreateFamilySizeTableManager(dstnctFamEdt, dstnctFamEdt.OwnerFamily.Id);
                    // If FSTM Created successfully then retrieve it
                    if (fstmResult)
                    {
                        //TaskDialog.Show("ImportSizeLookUpTable", "Starting with: " + "\n BIC: " + bic_nm + "\n dstnctFam.Name: " + dstnctFam.Name + "\n Getting an FSTM"); 
                        fstm = Autodesk.Revit.DB.FamilySizeTableManager.GetFamilySizeTableManager(dstnctFamEdt, dstnctFamEdt.OwnerFamily.Id);
                    }  // dstnctFamEdt.OwnerFamily.Id);
                }
                //TaskDialog.Show("ImportSizeLookUpTable", "Starting with: " + "\n BIC: " + bic_nm + "\n dstnctFam.Name: " + dstnctFam.Name + "\n dstnctFam.Id: " + dstnctFam.Id +
                //    "\n dstnctFamEdt.OwnerFamily.Id: " + dstnctFamEdt.OwnerFamily.Id + "\n FSTM: " + fstm);
                // If Retrieved FSTM is not Null, then use it.
                if (fstm != null)
                {
                    // Remove (Delete) existing File Before importing the new one.
                    using (Transaction removeSzLktbl_trx = new Transaction(dstnctFamEdt, "Importing csv File into Family "))
                    {
                        //TaskDialog.Show("ImportSizeLookUpTable", "Starting with: " + dstnctFam.Name);
                        removeSzLktbl_trx.Start("Start");

                        if (fstm.RemoveSizeTable(lookupFileName) == false)   // use filename only with extension and without Path
                        {
                            /*TaskDialog.Show("ImportSizeLookUpTable", "FamilyDocument: " + dstnctFamEdt.Title +
                                "\n- RemoveSizeTable Returned Null!! for file: " + csvFilePathName +
                                "\n- lookupFileName: " + lookupFileName);*/
                        }
                        // End transaction
                        removeSzLktbl_trx.Commit();
                        removeSzLktbl_trx.Dispose();
                        //TaskDialog.Show("ImportSizeLookUpTable", "after  Commit ");
                    }
                    // Import the csv file into Family.
                    using (Transaction importSzLktbl_trx = new Transaction(dstnctFamEdt, "Importing csv File into Family "))
                    {

                        //TaskDialog.Show("ImportSizeLookUpTable", "Starting with: " + dstnctFam.Name);
                        importSzLktbl_trx.Start("Start");

                        // Create Error Info  for ImportSizeTable
                        FamilySizeTableErrorInfo errorInfo = new FamilySizeTableErrorInfo();
                        //TaskDialog.Show("ImportSizeLookUpTable", "Created errorInfo ");

                        Autodesk.Revit.DB.FamilySizeTableErrorInfo ImportErrorInfo = new Autodesk.Revit.DB.FamilySizeTableErrorInfo();
                        if (fstm.ImportSizeTable(dstnctFamEdt, csvFilePathName, ImportErrorInfo) == false)
                        {    // Use Full path here
                            //TaskDialog.Show("ImportSizeLookUpTable", " ImportSizeTable Reteurned FALSE!!. " + "\n ImportErrorInfo: " + ImportErrorInfo.ToString());
                        }
                        fstm.Dispose();
                        // End transaction
                        importSzLktbl_trx.Commit();
                        //TaskDialog.Show("ImportSizeLookUpTable", "after  Commit ");
                    }// End Of using transaction importSzLktbl_trx
                    fstm.Dispose();
                }
                return importError;
            }
            catch (Exception ex)
            {
                //TaskDialog.Show("ImportSizeLookUpTable", " dstnctFam.Name: " + dstnctFamEdt.Title + "\n With CSV: " + csvFilePathName + "\n - lookuptableResult excepton: " + ex);
                return null;
            }
        }  // end of importSizeLookUpTable
        #endregion // End Of ImportSizeLookUpTable to Families

        #region   // Get_PipeFitting_FamilySymbols
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
        #endregion Get_PipeFitting_FamilySymbols

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
            string famDir = famManfDir + famName + "\\";                       // z:\BIM\KDS_TEMPLATE\2022\fittings\KDS_Char_CI_NH\KDS_Char_CI_NH-Coupling\

            string file = famDir + famName + ext;                        // z:\BIM\KDS_TEMPLATE\2022\fittings\KDS_Char_CI_NH\KDS_Char_CI_NH-Coupling\KDS_Char_CI_NH-Coupling.???

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
            else { selectedFileName = "Z:\\BIM\\Families\\SupplierCode\\KDS_CI_NH-BLK_MLBL.csv"; }
            openFileDialog1.Dispose();
            return selectedFileName;*/
            return "Z:\\BIM\\KDS_SUPPLIER_CODE\\KDS_CI_NH_BLK_MLBL.csv";
        }
        #endregion  // End Of Get Pricing Database filename

        #region   // Loop thru every Category and show Distinct Families   -- For Debug Only
        public void show_DstnctFam(List<bicFam_strct> bicDstnctFam_lst)
        {
            if (bicDstnctFam_lst.Count() > 0)
            {
                string ftf_tds = "bicDstnctFam.Count: " + bicDstnctFam_lst.Count();
                foreach (bicFam_strct bicDstnctFam in bicDstnctFam_lst)
                {
                    ftf_tds = "\n Family BIC: " + bicDstnctFam.bic_str + ":: Count: " + bicDstnctFam.fam_lst.Count;
                    if (bicDstnctFam.fam_lst.Count > 0)
                    {
                        int ftf_bic_int = 0;
                        foreach (Family dstnctFam in bicDstnctFam.fam_lst)
                        {
                            ftf_bic_int++;
                            ftf_tds += "\n" + ftf_bic_int + "- " + dstnctFam.Name;
                            TaskDialog.Show("show_DstnctFam", ftf_tds + "\n\n\n\n\n\n\n\n\n\n\n\n");
                        }  // for each dstnctFam
                    }  // if bicDstnctFam.fam_lst count > 0
                }  // foreach bicDstnctFam
            }// if bicFam_lst is >0
        }   // End OF show_DstnctFam
        #endregion   // End OF Loop thru every Category and Create the Dir Tree if it does not exist

        #region // bicFam_strct   It is a String to a List dictionary like, to hold all families of a certain BIC. so all ficture families in a bic_fixture_OST handle.
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

        #region   // Get external definition parameter (shared Param) by sting Name
        public List<ExternalDefinition> sort_listByAnother(List<ExternalDefinition> extDef_lst, List<string> paramFixedName_lst)
        {
            List<ExternalDefinition> extDef_Sorted_lst = new List<ExternalDefinition>();
            foreach (string fN in paramFixedName_lst)
            {
                foreach (ExternalDefinition extDef in extDef_lst)
                {
                    if (extDef.Name == fN)
                    {
                        extDef_Sorted_lst.Add(extDef);
                        continue;
                    }
                }
            }
            return extDef_Sorted_lst;
        }

        #endregion   //  End Of Get exteernal definition parameter (shared Param) by sting Name


        #region   // Hardcoded Family Parameter Formulas, So i can use later in code.  I am not sure if this is the smartest way, but considering its only 10 items, ithink it is ok.
        public List<famParam_Data_class> Get_HardCodedFamParamFormulas()
        {
            /// Hard Coded Formulas ///
            List<famParam_Data_class> fpData_lst = new List<famParam_Data_class>();

            famParam_Data_class fpData = new famParam_Data_class();
            fpData.Name = "KDS_ID_tbl";
            fpData.type = "other";
            fpData.units = "";
            fpData.dfltVlu = "abyz";
            fpData.formula = "";
            fpData_lst.Add(fpData);

            //TaskDialog.Show("sdfga", "\n fpData.Name: " + fpData.Name);
            famParam_Data_class fpData1 = new famParam_Data_class();

            fpData1.Name = "KDS_HPH";
            fpData1.type = "other";
            fpData1.units = "";
            fpData1.dfltVlu = "abyz";
            fpData1.formula = "size_lookup(KDS_ID_tbl,\"KDS_HPH\",\" \" ";   //, KDS_ND0, KDS_ND1, KDS_ND2, KDS_ND3)";
            fpData_lst.Add(fpData1);

            //TaskDialog.Show("sdfga", "\n fpData.Name: " + fpData1.Name);
            famParam_Data_class fpData2 = new famParam_Data_class();

            fpData2.Name = "KDS_MfrList";
            fpData2.type = "currency";
            fpData2.units = "currency";
            fpData2.dfltVlu = "999999";
            fpData2.formula = "size_lookup(KDS_ID_tbl, \"KDS_MfrList\", 999999 ";   //,KDS_ND0, KDS_ND1, KDS_ND2, KDS_ND3)";
            fpData_lst.Add(fpData2);

            //TaskDialog.Show("sdfga", "\n fpData.Name: " + fpData2.Name);
            famParam_Data_class fpData3 = new famParam_Data_class();

            fpData3.Name = "KDS_MfrPart";
            fpData3.type = "other";
            fpData3.units = "";
            fpData3.dfltVlu = "abyz";
            fpData3.formula = "size_lookup(KDS_ID_tbl, \"KDS_MfrPart\", \"abyz\" ";   //,KDS_ND0, KDS_ND1, KDS_ND2, KDS_ND3)";
            fpData_lst.Add(fpData3);

            //TaskDialog.Show("sdfga", "\n fpData.Name: " + fpData3.Name);
            famParam_Data_class fpData4 = new famParam_Data_class();

            fpData4.Name = "KDS_MCAA_LBR_RATE";
            fpData4.type = "number";
            fpData4.units = "general";
            fpData4.dfltVlu = "999999";
            fpData4.formula = "size_lookup(KDS_ID_tbl, \"KDS_MCAA_LBR_RATE\", 999999";   //,KDS_ND0, KDS_ND1, KDS_ND2, KDS_ND3)";
            fpData_lst.Add(fpData4);

            //TaskDialog.Show("sdfga", "\n fpData.Name: " + fpData4.Name);
            famParam_Data_class fpData5 = new famParam_Data_class();

            fpData5.Name = "KDS_LBR_RATE";
            fpData5.type = "number";
            fpData5.units = "general";
            fpData5.dfltVlu = "999999";
            fpData5.formula = "size_lookup(KDS_ID_tbl, \"KDS_LBR_RATE\", 999999 ";   //,KDS_ND0, KDS_ND1, KDS_ND2, KDS_ND3)";
            fpData_lst.Add(fpData5);

            //TaskDialog.Show("sdfga", "\n fpData.Name: " + fpData5.Name);
            famParam_Data_class fpData6 = new famParam_Data_class();

            fpData6.Name = "KDS_ND0";
            fpData6.type = "length";
            fpData6.units = "inches";
            fpData6.dfltVlu = "";
            fpData6.formula = "";
            fpData_lst.Add(fpData6);

            //TaskDialog.Show("sdfga", "\n fpData.Name: " + fpData6.Name);
            famParam_Data_class fpData7 = new famParam_Data_class();

            fpData7.Name = "KDS_ND1";
            fpData7.type = "length";
            fpData7.units = "inches";
            fpData7.dfltVlu = "";
            fpData7.formula = "";
            fpData_lst.Add(fpData7);

            //TaskDialog.Show("sdfga", "\n fpData.Name: " + fpData7.Name);
            famParam_Data_class fpData8 = new famParam_Data_class();

            fpData8.Name = "KDS_ND2";
            fpData8.type = "length";
            fpData8.units = "inches";
            fpData8.dfltVlu = "";
            fpData8.formula = "";
            fpData_lst.Add(fpData8);

            //TaskDialog.Show("sdfga", "\n fpData.Name: " + fpData8.Name);
            famParam_Data_class fpData9 = new famParam_Data_class();

            fpData9.Name = "KDS_ND3";
            fpData9.type = "length";
            fpData9.units = "inches";
            fpData9.dfltVlu = "";
            fpData9.formula = "";
            fpData_lst.Add(fpData9);

            /*string fpd_str = " Get_HardCodedFamParam List of Names.";
            foreach( famParam_Data_class fpD in fpData_lst )
             {
                 fpd_str = "\n fpData.Name: " + fpD.Name;
             }
             TaskDialog.Show("Get_HardCodedFamParamFormulas", "\n fpData.Name: " + fpd_str);
            */
            return fpData_lst;

        } // End Of Hardcoded Family Parameter Formulas
        #endregion

        #region   //Get an famParam_Data item from based on string name
        public famParam_Data_class get_famParam_Data_item(List<famParam_Data_class> famParam_Data_lst, string name)
        {
            famParam_Data_class get_famParam_Data_item = famParam_Data_lst.Where(n => n.Name == name).FirstOrDefault();
            //TaskDialog.Show("get_famParam_Data_item", "Name is: " + name + "\n item Name: " + get_famParam_Data_item.Name);
            return get_famParam_Data_item;
        }// End Of get_famParam_Data_item
        #endregion   // End Of Get an famParam_Data item from based on string name

        #region   //Get an famParam_Data item from based on string name
        public string get_famParam_Data_Str(List<famParam_Data_class> famParam_Data_lst, string name)
        {
            famParam_Data_class get_famParam_Data_item = famParam_Data_lst.Where(n => n.Name == name).FirstOrDefault();
            //TaskDialog.Show("get_famParam_Data_item", "Name is: " + name + "\n item Name: " + get_famParam_Data_item.Name);
            return get_famParam_Data_item.formula;
        }// End Of get_famParam_Data_item
        #endregion   // End Of Get an famParam_Data item from based on string name


        #region  // get Parameter by name... since GetParameters returns a list.... so this could be modified to do more checking.
        public FamilyParameter getParmeter_byName(FamilyManager famMan, string paramName)
        {
            foreach (FamilyParameter famParam in famMan.GetParameters())
            {
                if (famParam.Definition.Name == paramName)
                {
                    return famParam;
                }
            }
            return null;
        }  // End Of getParmeter_byName

        #endregion    // End Of get Parameter by name

        
        #region   // Get Associated Parameter for a given connectorElement
        public string get_AssociatedParametersofConnectorElement(Document dstnctFamDocEdt, ConnectorElement connectorElement)
        {
            try
            {
                if (connectorElement != null)
                {
                    var radiusPara = connectorElement.get_Parameter(BuiltInParameter.CONNECTOR_RADIUS);
                    foreach (FamilyParameter familyPara in dstnctFamDocEdt.FamilyManager.Parameters)
                    {
                        foreach (Parameter associatedPara in familyPara.AssociatedParameters)
                        {
                            if (radiusPara.Id == associatedPara.Id && associatedPara.Element.Id == connectorElement.Id)
                            {
                                //associate parameter found
                                return familyPara.Definition.Name;
                            }
                        }
                    }
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
            }
            return null;
        }
        #endregion  // End Of Get Associated Parameter for a given connectorElement


        //======================================== create_update_Fam_KDS_CSV ===========================================================//



    }  //end of Class create_update_Fam_KDS_CSV
    #endregion  // create_update_Fam_KDS_CSV


    //======================================== EST_DATA_CLASS ===========================================================//







   





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