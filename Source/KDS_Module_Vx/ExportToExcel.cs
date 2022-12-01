using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Utility;   // For Common Classes and Functions 



namespace KDS_Module_Vx
{
    public class ExportToExcel : IExternalCommand
    {


        #region // Execute Region of Code (Invoked by KDS Ribbon) //
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document actvDoc = uiDoc.Document;
            //Autodesk.Revit.DB.Category cati = null; // new Autodesk.Revit.DB.Category("BuiltInCategory.OST_PipeFitting");
            //CreateSharedParameter(actvDoc, cati, 1, ParameterType.Currency, false);   / Not needed it was for testing creating a sharedParameter.
            //return Result.Succeeded;

            // Lists for Headers in Excel File. Divided into categories based on type of parameter //
            List<string> elm_hdr_lst = new List<string> { "ElementID", "Family", "Type" };
            List<string> fam_hdr_lst = new List<string> { "System Classification", "KDS_MCAA_LBR_RATE", "KDS_LBR_RATE", "KDS_HPH", "KDS_MfrList", "KDS_MfrPart", "Category", "Size", "Length", "System Type" };
            List<string> calc_hdr_lst = new List<string> { "Is Vertical", "Level", "Diameter", "System Name" };


            // Get Alll Pipe KDS_Est_Data into a DB class
            const string KDS_EST_PipeData_CSV_path = "Z:\\BIM\\KDS_SUPPLIER_CODE\\";
            const string KDS_EST_PipeData_CSV_fn = "KDS_All_Pipes.csv";
            string KDS_EST_PipeData_CSV = KDS_EST_PipeData_CSV_path + KDS_EST_PipeData_CSV_fn;
            // DEbug Test reading Pipe Data File into a DB Class
            List<est_data_class> KDSEstData_Pipe_lst = get_KDSEstData_Pipe(KDS_EST_PipeData_CSV);
            //return Result.Succeeded;


            #region  // Get Embedded dll.  I was able to do without this using nuget for DocumentFormat-OpenXML.dll. It is here temp in case i still need it.
            //Calling of functions //
            //LoadReference();
            #endregion // End Of Get Embedded dll.  I was able to do without this using nuget for DocumentFormat-OpenXML.dll. It is here temp in case i still need it.


            // Get all Family Instances
            //List<FamilyInstance> famInst_lst = new FilteredElementCollector(uiDoc.Document).OfCategory(bic).WhereElementIsNotElementType().Select(df => df as FamilyInstance).ToList();

            List<FamilyInstance> famInst_lst = GetPipeFittings(uiDoc); // new FilteredElementCollector(uiDoc.Document).WhereElementIsNotElementType().Select(df => df as FamilyInstance).ToList();


            /* p.Id.ToString(), p.Name, p.GetType().Name, p.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsString(),
                         "\"NA\"", "\"NA\"", "\"NA\"", "\"NA\"", "\"NA\"", p.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString(),
                         p.LookupParameter("Size").AsString() , p.LookupParameter("Length").AsValueString() , p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString(),
                         IsVertical(p), GetPipeLevel(uiDoc, p, "middle"), (12 * p.Diameter).ToString() + "\"", p.MEPSystem.Name,
 */


            if (famInst_lst.Count > 0)
            {
                string tmpStr = "";
                int tmpCnt = 0;

                foreach (FamilyInstance fi in famInst_lst)
                {
                    tmpCnt++;
                    string tempSys = fi.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsString();
                    //FamilyParameter tempParam = fi.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM);


                    //if (tempSys.Contains(',')) { tempSys = tempSys.Split(',').First(); }
                    tmpStr += tmpCnt + "- " + tempSys +
                        " :: " + fi.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString() +
                        " :: " + fi.Name + "\n";
                }
                // Group By BIC
                //List <string, FamilyInstance> BICFamInst_lst = famInst_lst.GroupBy(fi => fi.Category);    //how to find unique families based on family name, not family type, of given categories

                // Group By System

                List<FamilyInstance> SysBICFamInst_lst = famInst_lst.GroupBy(f => f.GetType()).Select(Sys => Sys.First()).ToList();    //how to find unique families based on family name, not family type, of given categories

                //var xyz =  SysBICFamInst_lst.GroupBy(x=> new {x.Category,x.GetType().Name},(key, group) => new { Key1 = key.Category, Key2 = key.GetType().Name, Result = group.ToList() });
                tmpStr += "\n\n\n\n\n\n\n\n\n\n\n\n";
                //TaskDialog.Show("ExportToExcel", tmpStr);
            }
            else
            {
                TaskDialog.Show("ExportToExcel", " Count 0");
            }





            #region // Dialog Box to get input from user. Unused now, but left here as model if needed.
            bool dialogBox_rslt = true;

            dialogBox_rslt = dialogBox(uiDoc, elm_hdr_lst, fam_hdr_lst, calc_hdr_lst, KDSEstData_Pipe_lst);
            if (dialogBox_rslt)
            {
                exportToExcel_func(uiDoc, elm_hdr_lst, fam_hdr_lst, calc_hdr_lst, KDSEstData_Pipe_lst);

                //List<List<string>> rc_data = exportToExcel_func(uiDoc, elm_hdr_lst, fam_hdr_lst, calc_hdr_lst,plumElem_lst);

                // Write Family and Fixture Data to arc_data to a Macro Enable excel (.xlsm)
                //              CreateXL_macro(filePath, sheetName, startRow, startCol, rc_data, hdrList);

            }
            else
            {
                TaskDialog.Show("", "Nothing to Do.");
            }

            #endregion // End Of Dialog Box to get input from user. Unused now, but left here as model if needed.

            return Result.Succeeded;
        }  // End Of Execute
        #endregion



        #region // Load Reference Function //  unused, part of loading DocumentFormat.OpenXML.dll
        [STAThread]
        static void LoadReference()
        {
            string resource1 = "KDS_Revit_Commands.DocumentFormat.OpenXml.dll"; // ClassName.EmbeddedResource
            //string resource1_loc = "C:\\Users\\KDS-EST-3\\source\\repos\\KDS_Module_Vx\\Source\\packages\\Open-XML-SDK.2.9.1\\lib\\net46\\DocumentFormat.OpenXml.dll";
            string resource1_loc = "C:\\Users\\KDS-EST-3\\source\\repos\\KDS_Module_Vx\\Source\\packages\\DocumentFormat.OpenXmlSDK.2.0\\lib\\Net35\\DocumentFormat.OpenXml.dll";
            EmbeddedAssembly.Load(resource1, resource1_loc);     // Raises Event

            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve); // Event Handler
        }

        static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return EmbeddedAssembly.Get(args.Name); // event, gets embedded resource
        }
        #endregion  // End Of Load Reference Function //  unused, part of loading DocumentFormat.OpenXML.dll 

        #region // DialogBox Function to handle Facilitation of Function //
        public bool dialogBox(UIDocument uidoc, List<string> elmHdrList, List<string> famHdrList, List<string> calcHdrList, List<est_data_class> KDSEstData_Pipe_lst)
        {
            #region // DialogBox Settings 
            TaskDialog mainDialog = new TaskDialog("Export to Excel")
            {
                MainInstruction = "Please choose an option",
                MainContent = "Choose an option below",
                Title = "Export To Excel",
                CommonButtons = TaskDialogCommonButtons.Close,
                DefaultButton = TaskDialogResult.Close,
                FooterText = "KDS Plumbing and Heating Services",
                MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                TitleAutoPrefix = false,
            };

            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Export to Excel");

            TaskDialogResult tResult = mainDialog.Show(); // returns the users response from maindialog

            #endregion // DialogBox settings

            #region // Set up of Logic to respond to user's main dialog choice //

            if (TaskDialogResult.CommandLink1 == tResult)
            { return true; }
            else
            { return false; }
            #endregion


        } // end DialogBox Macro
        #endregion


        #region // exportToExcel_func_by_List    DO NOT USE.. this is meant to have all instances sorted out into one large List of lists by system name.
        /*
                public List<List<string>> exportToExcel_func(UIDocument uiDoc, List<string> elm_hdr_lst, List<string> famHdrList, List<string> calcHdrList, List<FamilyInstance> plumElem_lst)
                {
                    #region // User Chose Export to Excel Option //

                    // Collection of Plumbing Elements Lists //
                    List<Pipe> pipes = GetPipes(uiDoc);
                    List<FamilyInstance> fittings_lst = GetPipeFittings(uiDoc);
                    List<FamilyInstance> fixtures_lst = GetPlumbingFixtures(uiDoc);
                    List<FamilyInstance> pipeAcce_lst = GetPlumbingPipeAccessories(uiDoc);

                    // creation of header titles for Excel File //
                    List<string> hdrList = new List<string>();
                    hdrList.AddRange(elm_hdr_lst);
                    hdrList.AddRange(famHdrList);
                    hdrList.AddRange(calcHdrList);

                    // Initialization for Excel Sheet Template to Write Our Data to//
                    string filePath = "C:\\Users\\KDS-EST-3\\Desktop\\Excel_Files\\Elements.xltm";
                    string sheetName = "importedData";
                    int startRow = 1;
                    int startCol = 1;

                    #region // Adding of data to rc_data list //
                    List<List<string>> rc_data = new List<List<string>>();

                    // Adding Pipe data to rc_data
                    foreach (Pipe pipe in pipes)
                    {
                        rc_data.Add(new List<string>
                            {
                            pipe.Id.ToString(),
                            pipe.Name,
                            pipe.GetType().Name,
                            pipe.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsString(),
                            get_Pipe_KDSParams(pipe,"KDS_MCAA_LBR_RATE"),
                            get_Pipe_KDSParams(pipe,"KDS_LBR_RATE"),
                            get_Pipe_KDSParams(pipe,"KDS_HPH"),
                            get_Pipe_KDSParams(pipe,"KDS_MfrList"),
                            get_Pipe_KDSParams(pipe,"KDS_MfrPart"),
                            pipe.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString(),
                            pipe.LookupParameter("Size").AsString() , pipe.LookupParameter("Length").AsValueString(),
                            pipe.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString(),
                            IsVertical(pipe),
                            GetPipeLevel(uiDoc, pipe, "middle"),
                            (12 * pipe.Diameter).ToString() + "\"",
                            pipe.MEPSystem.Name,
                            }); 
                    }

                    // Adding Pipe Fitting Data to rc_data
                    foreach (FamilyInstance Fitt in fittings_lst)
                    {
                        string mcaa_lbr_rate = Fitt?.LookupParameter("KDS_MCAA_LBR_RATE")?.AsValueString() ?? "\"NA\"";
                        string lbr_rate = Fitt?.LookupParameter("KDS_LBR_RATE")?.AsValueString() ?? "\"NA\"";
                        string hph = Fitt?.LookupParameter("KDS_HPH")?.AsString() ?? "\"NA\"";
                        string mfrList = Fitt?.LookupParameter("KDS_MfrList")?.AsValueString() ?? "\"NA\"";
                        string mfrPart = Fitt?.LookupParameter("KDS_MfrPart")?.AsString() ?? "\"NA\"";
                        string size = Fitt?.LookupParameter("Size")?.AsString() ?? "\"NA\"";

                        rc_data.Add(new List<string>
                        {
                            Fitt.Id.ToString(),
                            Fitt.Symbol.Name,
                            Fitt.GetType().Name,
                            Fitt.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsString(),
                            mcaa_lbr_rate,
                            lbr_rate,
                            hph,
                            mfrList,
                            mfrPart,
                            Fitt.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString(),
                            size,
                            "\"NA\"",
                            Fitt.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString(),
                            "\"NA\"",
                            GetInstanceLevel(uiDoc, Fitt),
                            "\"NA\"",
                            Fitt.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString(),
                            });
                    }

                    // Adding Plumbing Fixture data to rc_data
                    foreach (FamilyInstance fixt in fixtures_lst)
                    {
                        rc_data.Add(new List<string>
                        {
                            fixt.Id.ToString(),
                            fixt.Symbol.Name,
                            fixt.GetType().Name,
                            fixt.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsString(),
                            "\"NA\"",
                            "\"NA\"",
                            "\"NA\"",
                            "\"NA\"",
                            "\"NA\"",
                            fixt.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString(),
                            "\"NA\"",
                            "\"NA\"" ,
                            fixt.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString(),
                            "\"NA\"",
                            GetInstanceLevel(uiDoc, fixt),
                            "\"NA\"",
                            fixt.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString(),
                            });
                    }

                    // Adding Pipe Accessories data to rc_data
                    foreach (FamilyInstance PipeAcce in pipeAcce_lst)
                    {
                        rc_data.Add(new List<string>
                        {
                            PipeAcce.Id.ToString(),
                            PipeAcce.Symbol.Name,
                            PipeAcce.GetType().Name,
                            PipeAcce.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsString(),
                            "\"NA\"",
                            "\"NA\"",
                            "\"NA\"",
                            "\"NA\"",
                            "\"NA\"",
                            PipeAcce.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString(),
                            "\"NA\"" ,
                            "\"NA\"" ,
                            PipeAcce.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString(),
                            "\"NA\"",
                            GetInstanceLevel(uiDoc, PipeAcce),
                            "\"NA\"",
                            PipeAcce.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString(),
                            });
                    }

                    return rc_data;

                    #endregion


                    #endregion

                }  // End of exportToExcel_func
        */
        #endregion  // End Of exportToExcel_func




        #region // exportToExcel_func   
        public void exportToExcel_func(UIDocument uiDoc, List<string> elm_hdr_lst, List<string> famHdrList, List<string> calcHdrList, List<est_data_class> KDSEstData_Pipe_lst)
        {
            // User Chose Export to Excel Option  //
            #region  //  Initialization and Definitions

            // Declare and Define Collection of Plumbing Elements Lists //
            List<Pipe> pipes = GetPipes(uiDoc);
            List<FamilyInstance> fittings_lst = GetPipeFittings(uiDoc);
            List<FamilyInstance> fixtures_lst = GetPlumbingFixtures(uiDoc);
            List<FamilyInstance> pipeAcce_lst = GetPlumbingPipeAccessories(uiDoc);

            // creation of header titles for Excel File //
            List<string> hdrList = new List<string>();
            hdrList.AddRange(elm_hdr_lst);   // Header Names 
            hdrList.AddRange(famHdrList);    // List of family instances
            hdrList.AddRange(calcHdrList);   // List of Calculated Values such as is Vertical or IsUnderground etc

            // Initialization of Output Excel Sheet Template to Write Our Data to//
            string filePath = "C:\\Users\\KDS-EST-3\\Desktop\\Excel_Files\\Elements.xltm";   //// move this to be an input param to this function.//
            string sheetName = "importedData";
            int startRow = 1;
            int startCol = 1;
            #endregion  //  End Of Initialization and Definitions


            #region // Adding of data to rc_data list //  rc_data is a list Of Lists that will hold all Pipes,Fittings and Accessories INSTANCES found in the project
            List<List<string>> rc_data = new List<List<string>>();

            // Adding Pipe data to rc_data
            foreach (Pipe pipe in pipes) 
            {
                // I left these defined here for debugging purpose.  i will roll into rc_data.Add directly to save on creating them in every loop
            
                string id = pipe.Id.ToString();
                string name = pipe.Name;
                string getTypeName = pipe.GetType().Name;
                string getSysClassName = ""; if (pipe.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM) != null) { getSysClassName = pipe.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsString(); }

                string id1 = get_Pipe_KDSParams(KDSEstData_Pipe_lst, pipe, "KDS_MCAA_LBR_RATE");
                string id2 = get_Pipe_KDSParams(KDSEstData_Pipe_lst, pipe, "KDS_LBR_RATE");
                string id3 = get_Pipe_KDSParams(KDSEstData_Pipe_lst, pipe, "KDS_HPH");
                string id4 = get_Pipe_KDSParams(KDSEstData_Pipe_lst, pipe, "KDS_MfrList");
                string id5 = get_Pipe_KDSParams(KDSEstData_Pipe_lst, pipe, "KDS_MfrPart");

                string getElemCatName = pipe?.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM)?.AsValueString();
                string size = pipe?.LookupParameter("Size")?.AsString();
                string length = pipe?.LookupParameter("Length")?.AsDouble().ToString("0.000");   // AsValueString();
                string getPipeSysTypeName = pipe?.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM)?.AsValueString();
                string isVertical = IsVertical(pipe);
                string getLevel = GetPipeLevel(uiDoc, pipe, "middle");
                string diam = (12 * pipe.Diameter).ToString() + "\"";
                string MEPSysName = pipe?.MEPSystem?.Name != null ? pipe.MEPSystem.Name : "null";

                rc_data.Add(new List<string>
                {
                    id,
                    name,
                    getTypeName,
                    getSysClassName,
                    id1,
                    id2,
                    id3,
                    id4,
                    id5,
                    getElemCatName,
                    size,
                    length,
                    getPipeSysTypeName,
                    isVertical,
                    getLevel,
                    diam,
                    MEPSysName,
                }
                );
            }

            #region // Adding Pipe Fitting Data to rc_data
            foreach (FamilyInstance Fitt in fittings_lst)
            {
                /*
                string mcaa_lbr_rate = Fitt?.LookupParameter("KDS_MCAA_LBR_RATE")?.AsValueString() ?? "\"NA\"";
                string lbr_rate = Fitt?.LookupParameter("KDS_LBR_RATE")?.AsValueString() ?? "\"NA\"";
                string hph = Fitt?.LookupParameter("KDS_HPH")?.AsString() ?? "\"NA\"";
                string mfrList = Fitt?.LookupParameter("KDS_MfrList")?.AsValueString() ?? "\"NA\"";
                string mfrPart = Fitt?.LookupParameter("KDS_MfrPart")?.AsString() ?? "\"NA\"";
                string size = Fitt?.LookupParameter("Size")?.AsString() ?? "\"NA\"";
*/
                rc_data.Add(new List<string>
                {
                    Fitt.Id.ToString(),
                    Fitt.Symbol.Name,
                    Fitt.GetType().Name,
                    Fitt.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsString(),
                    Fitt?.LookupParameter("KDS_MCAA_LBR_RATE")?.AsValueString() ?? "\"NA\"",
                    Fitt?.LookupParameter("KDS_LBR_RATE")?.AsValueString() ?? "\"NA\"",
                    Fitt?.LookupParameter("KDS_HPH")?.AsString() ?? "\"NA\"",
                    Fitt?.LookupParameter("KDS_MfrList")?.AsValueString() ?? "\"NA\"",
                    Fitt?.LookupParameter("KDS_MfrPart")?.AsString() ?? "\"NA\"",
                    Fitt.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString(),
                    Fitt?.LookupParameter("Size")?.AsString() ?? "\"NA\"",
                    "\"NA\"",
                    Fitt.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString(),
                    "\"NA\"",
                    GetInstanceLevel(uiDoc, Fitt),
                    "\"NA\"",
                    Fitt.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString(),
                    });
            }
            #endregion // End Of Adding Pipe Fitting Data to rc_data


            #region // Adding Plumbing Fixture data to rc_data
            foreach (FamilyInstance fixt in fixtures_lst)
            {
                rc_data.Add(new List<string>
                {
                    fixt.Id.ToString(),
                    fixt.Symbol.Name,
                    fixt.GetType().Name,
                    fixt.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsString(),
                    fixt?.LookupParameter("KDS_MCAA_LBR_RATE")?.AsValueString() ?? "\"NA\"",
                    fixt?.LookupParameter("KDS_LBR_RATE")?.AsValueString() ?? "\"NA\"",
                    fixt?.LookupParameter("KDS_HPH")?.AsString() ?? "\"NA\"",
                    fixt?.LookupParameter("KDS_MfrList")?.AsValueString() ?? "\"NA\"",
                    fixt?.LookupParameter("KDS_MfrPart")?.AsString() ?? "\"NA\"",
                    fixt.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString(),
                    "\"NA\"",
                    "\"NA\"" ,
                    fixt.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString(),
                    "\"NA\"",
                    GetInstanceLevel(uiDoc, fixt),
                    "\"NA\"",
                    fixt.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString(),
                    });
            }   // foreach fixt
            #endregion // Adding Plumbing Fixture data to rc_data


            #region // Adding Pipe Accessories data to rc_data
            foreach (FamilyInstance PipeAcce in pipeAcce_lst)
            {
                rc_data.Add(new List<string>
                {
                    PipeAcce.Id.ToString(),
                    PipeAcce.Symbol.Name,
                    PipeAcce.GetType().Name,
                    PipeAcce.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsString(),
                    PipeAcce?.LookupParameter("KDS_MCAA_LBR_RATE")?.AsValueString() ?? "\"NA\"",
                    PipeAcce?.LookupParameter("KDS_LBR_RATE")?.AsValueString() ?? "\"NA\"",
                    PipeAcce?.LookupParameter("KDS_HPH")?.AsString() ?? "\"NA\"",
                    PipeAcce?.LookupParameter("KDS_MfrList")?.AsValueString() ?? "\"NA\"",
                    PipeAcce?.LookupParameter("KDS_MfrPart")?.AsString() ?? "\"NA\"",
                    PipeAcce.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString(),
                    "\"NA\"" ,
                    "\"NA\"" ,
                    PipeAcce.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString(),
                    "\"NA\"",
                    GetInstanceLevel(uiDoc, PipeAcce),
                    "\"NA\"",
                    PipeAcce.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString(),
                    });
            }  // End of foreach PipeAcce
            #endregion  // End Of Adding Pipe Accessories data to rc_data

            #endregion
            // Write Family and Fixture Data to arc_data to a Macro Enable excel (.xlsm)
            CreateXL_macro(filePath, sheetName, startRow, startCol, rc_data, hdrList);


            

        }  // End of exportToExcel_func
        #endregion  // End Of exportToExcel_func



        #region // GetPipes Function to collect all Pipe Elements //
        public static List<Pipe> GetPipes(UIDocument uidoc)
        {
            Autodesk.Revit.DB.Document actvDoc = uidoc.Document;
            List<Pipe> elements = new List<Pipe>();

            FilteredElementCollector pipeCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(Pipe));

            foreach (Pipe p in pipeCollector)
            {
                elements.Add(p);
            }
            return elements;

            /*
            // Get a List of All Pipes that are Vertical and within 2" and 6" size
            FilteredElementCollector pipeCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(Pipe));
            List<Pipe> pipes_lst = pipeCollector.Cast<Pipe>().ToList();

            Parameter p_SysType = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);



            FilteredElementCollector instCollector = new FilteredElementCollector(actvDoc);

            List<FamilyInstance> allFamInst = instCollector.OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                                                                                            .Where(inst => inst.Symbol.Id.Equals(currFamSymbol.Id) &&
                                                                                                   inst.LookupParameter(param_name).AsDouble() >= lowerAngleLimit &&
                                                                                                   inst.LookupParameter(param_name).AsDouble() <= upperAngleLimit)
                                                                                            .ToList();
*/
        }
        #endregion



       
        #region   // get_Pipe_KDSParams    Returns the parameter value of a KDS Parameter from KDS_pipe Supplier files
        // I need a file for pipes only that i will read anf get this info from.
        //This file should be always available with each run.
        // See if it is possible or easier to deal with if i turn it into a fitting, or add it as a shared parameter file.
        // if shared parameter, it may be easier to read and write to it??
        public string get_Pipe_KDSParams(List<est_data_class> KDSEstData_Pipe_lst, Pipe pipe, string kdsParamName)
        {
            string name = pipe.Name;
            string mcaa_lbr_rate = pipe?.LookupParameter("KDS_MCAA_LBR_RATE")?.AsValueString() ?? "\"NA\"";
            string lbr_rate = pipe?.LookupParameter("KDS_LBR_RATE")?.AsValueString() ?? "\"NA\"";
            string hph = pipe?.LookupParameter("KDS_HPH")?.AsString() ?? "\"NA\"";
            string mfrList = pipe?.LookupParameter("KDS_MfrList")?.AsValueString() ?? "\"NA\"";
            string mfrPart = pipe?.LookupParameter("KDS_MfrPart")?.AsString() ?? "\"NA\"";
            string size = pipe?.LookupParameter("Size")?.AsString() ?? "\"NA\"";
            
            est_data_class result = KDSEstData_Pipe_lst.Find(edc => edc.famName == pipe.Name && edc.Size == size.Remove(size.Length-1,1));


           
            return result?[kdsParamName]?.ToString()?? "\"NA\"";

        }   // End Of get_Pipe_KDSParams
        #endregion  // End Of get_Pipe_KDSParams    Returns the parameter value of a KDS Parameter from KDS_pipe Supplier files



        #region // GetPipeFittings Function to collect all Pipe Fittings //
        public static List<FamilyInstance> GetPipeFittings(UIDocument uiDoc)
        {
            Autodesk.Revit.DB.Document actvDoc = uiDoc.Document;
            List<FamilyInstance> elements = new List<FamilyInstance>();

            FilteredElementCollector fittingCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_PipeFitting);

            foreach (FamilyInstance fitt in fittingCollector)
            {
                elements.Add(fitt);
            }

            return elements;
        }  // End Of GetPipeFittings
        #endregion   // End Of GetPipeFittings Function to collect all Pipe Fittings //

        #region // GetPlumbingFixtures Function to collect all Plumbing Fixtures //
        public static List<FamilyInstance> GetPlumbingFixtures(UIDocument uiDoc)
        {
            Autodesk.Revit.DB.Document actvDoc = uiDoc.Document;
            List<FamilyInstance> elements = new List<FamilyInstance>();

            FilteredElementCollector fixtureCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_PlumbingFixtures);

            foreach (FamilyInstance fixt in fixtureCollector)
            {
                elements.Add(fixt);
            }
            return elements;
        }   // End Of GetPlumbingFixtures
        #endregion  // End Of GetPlumbingFixtures Function to collect all Plumbing Fixtures //



        #region // GetPlumbingPipeAccessories Function to collect all Plumbing Pipe Accessories (Hammer Arresstor, Sleeves) //
        public static List<FamilyInstance> GetPlumbingPipeAccessories(UIDocument uiDoc)
        {
            Autodesk.Revit.DB.Document actvDoc = uiDoc.Document;
            List<FamilyInstance> elements = new List<FamilyInstance>();

            FilteredElementCollector PipeAccesCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_PipeAccessory);

            foreach (FamilyInstance pa in PipeAccesCollector)
            {
                elements.Add(pa);
            }
            return elements;
        }  // End of GetPlumbingPipeAccessories
        #endregion  // End Of GetPlumbingPipeAccessories




        #region // Get_DocLevel_strings Function to collect Names of Levels //
        public static List<string> Get_DocLevel_strings(UIDocument uidoc)
        {
            Autodesk.Revit.DB.Document actvDoc = uidoc.Document;

            List<string> docLevels_Names_lst = new FilteredElementCollector(actvDoc)
                                                                     .OfClass(typeof(Autodesk.Revit.DB.Level))
                                                                     .Select(e => e.Name)
                                                                     .Where(e => e != null)
                                                                     .ToList();
            return docLevels_Names_lst;
        }
        #endregion

        #region // GetPipeLevel Function to get Level Name of Current Pipe //
        public string GetPipeLevel(UIDocument uidoc, Pipe p, string startEndMiddle)
        {
            string lvlName = null;
            LocationCurve lc = p.Location as LocationCurve;
            Curve c = lc.Curve;
            double zLocStart = c.GetEndPoint(0).Z;
            double zLocEnd = c.GetEndPoint(1).Z;
            double zLocMiddle = (zLocStart + zLocEnd) / 2;

            Autodesk.Revit.DB.Document actvDoc = uidoc.Document;

            List<Autodesk.Revit.DB.Level> levels = new List<Autodesk.Revit.DB.Level>();
            List<Autodesk.Revit.DB.Level> sortedLevels = new List<Autodesk.Revit.DB.Level>();

            FilteredElementCollector lvlCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(Autodesk.Revit.DB.Level));

            foreach (Autodesk.Revit.DB.Level level in lvlCollector)
            {
                levels.Add(level);
            }

            sortedLevels = levels.OrderBy(o => o.Elevation).ToList();

            #region // User wants Z-Coordinate of Start EndPoint for Pipe //

            if (startEndMiddle == "start" || startEndMiddle == "Start")
            {
                foreach (Autodesk.Revit.DB.Level level in sortedLevels)
                {
                    if (zLocStart > level.Elevation)
                    {
                        lvlName = level.Name;
                    }
                }

                if (lvlName == null)
                {
                    lvlName = sortedLevels.First().Name;
                }
            }
            #endregion

            #region // User wants Z-Coordinate of End Endpoint for Pipe //
            if (startEndMiddle == "end" || startEndMiddle == "End")
            {
                foreach (Autodesk.Revit.DB.Level level in sortedLevels)
                {
                    if (zLocEnd > level.Elevation)
                    {
                        lvlName = level.Name;
                    }
                }

                if (lvlName == null)
                {
                    lvlName = sortedLevels.First().Name;
                }
            }
            #endregion

            #region // User wants Z-Coordinate of Midpoint of Pipe //
            if (startEndMiddle == "middle" || startEndMiddle == "Middle")
            {
                foreach (Autodesk.Revit.DB.Level level in sortedLevels)
                {
                    if (zLocMiddle > level.Elevation)
                    {
                        lvlName = level.Name;
                    }
                }

                if (lvlName == null)
                {
                    lvlName = sortedLevels.First().Name;
                }
            }
            #endregion

            else { TaskDialog.Show("error", "value for startEndMiddle is not \"start\", \"end\", or \"middle\""); }

            return lvlName;
        }
        #endregion

        #region // GetFittingLevel Function to get Level Name of Current Fitting //
        public string GetInstanceLevel(UIDocument uidoc, FamilyInstance inst)
        {
            Autodesk.Revit.DB.Document actvDoc = uidoc.Document;

            string lvlName = null;
            LocationPoint lp = inst.Location as LocationPoint;
            double zLoc = lp.Point.Z;

            List<Autodesk.Revit.DB.Level> levels = new List<Autodesk.Revit.DB.Level>();
            List<Autodesk.Revit.DB.Level> sortedLevels = new List<Autodesk.Revit.DB.Level>();

            FilteredElementCollector lvlCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(Autodesk.Revit.DB.Level));

            foreach (Autodesk.Revit.DB.Level level in lvlCollector)
            {
                levels.Add(level);
            }

            sortedLevels = levels.OrderBy(o => o.Elevation).ToList();

            foreach (Autodesk.Revit.DB.Level level in sortedLevels)
            {
                if (zLoc > level.Elevation)
                {
                    lvlName = level.Name;
                }
            }

            if (lvlName == null)
            {
                lvlName = sortedLevels.First().Name;
            }

            return lvlName;
        }
        #endregion

        #region // IsVertical Function to determine whether a Pipe is vertical or not //
        public string IsVertical(Pipe p)
        {
            double tolerance = 0.01;
            double bottom_elv = p.get_Parameter(BuiltInParameter.RBS_PIPE_BOTTOM_ELEVATION).AsDouble();
            double top_elv = p.get_Parameter(BuiltInParameter.RBS_PIPE_TOP_ELEVATION).AsDouble();
            double length = p.LookupParameter("Length").AsDouble();
            double calcLength, lenDiff_perc;

            if (bottom_elv * top_elv * length != 0)
            {
                calcLength = Math.Abs(top_elv - bottom_elv);
                lenDiff_perc = Math.Abs((length - calcLength) / length);

                if (lenDiff_perc <= tolerance)
                {
                    return "Yes";
                }
                else
                {
                    return "No";
                }
            }
            else
            {
                return "NA";
            }
        }
        #endregion

        #region // CreateXL_macro Function to Handle Exporting (Writing) rc_data to a Macro Enable excel (.xlsm)
        public static void CreateXL_macro(string filePath, string sheetName, int startRow, int startCol, List<List<string>> rc_data, List<string> hdrList)
        {
            #region // Check if FilePath Exists if not open filedialog foe user to choose a template file
            if (File.Exists(filePath))
            {
                //file exist
            }
            else
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "C:\\";
                    openFileDialog.Filter = "Excel_files (*.xltm)|*.xltm|All files (*.*)|*.*";
                    openFileDialog.FilterIndex = 2;
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //Get the path of specified file
                        filePath = openFileDialog.FileName;
                    }
                }
            }
            #endregion // Check if FilePath Exists if not open filedialog foe user to choose a template file

            //  Load the file into a byteArry to transfer to MemoryStream Later
            byte[] xltmDoc_byteArray = File.ReadAllBytes(filePath);
            using (MemoryStream xltmDoc_mem = new MemoryStream())
            {
                xltmDoc_mem.Write(xltmDoc_byteArray, 0, xltmDoc_byteArray.Length);

                using (SpreadsheetDocument xltmDoc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(xltmDoc_mem, true))
                {
                    xltmDoc.ChangeDocumentType(SpreadsheetDocumentType.MacroEnabledWorkbook);

                    #region  // Writing to Excel

                    WriteExcelFile(xltmDoc, sheetName, startRow, startCol, rc_data, hdrList);

                    #endregion  //  Writing to Excel

                    #region  // Set filepath to unique value

                    filePath = filePath.Replace("xltm", "xlsm");

                    // Check if filePath does not exist and increment name if it does.
                    filePath = AppendFileNumberIfExists(filePath, ".xlsm");

                    #endregion  // Set filepath to unique value

                }  // END using SpreadSheet

                #region  //  Change File Type and Save (XLTM to XLSM)
                // At this point, the memory stream contains the modified document.
                // We could write it back to a SharePoint document library or serve
                // it from a web server.
                // In this example, we serialize back to the file system to verify
                // that the code worked properly.
                using (FileStream fileStream = new FileStream(filePath, System.IO.FileMode.CreateNew))
                {
                    xltmDoc_mem.WriteTo(fileStream);

                }  // END using FileStream
                #endregion  //  Change File Type and Save (XLTM to XLSM)

            }  // END using MemoryStream
        }  // END of createXL
        #endregion

        #region // WriteExcelFile Function to write to data to excel file //
        private static void WriteExcelFile(SpreadsheetDocument xltmDoc, string sheetName, int startRow, int startCol, List<List<string>> rc_data, List<string> hdrList)
        {
            using (xltmDoc)
            {
                // Get the SharedStringTablePart. If it does not exist, create a new one.
                SharedStringTablePart shareStringPart;
                if (xltmDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = xltmDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = xltmDoc.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }

                #region  // Find desired worksheet by name or create one if it does not exist
                WorksheetPart worksheetPart = GetWorksheetPartByName(xltmDoc, sheetName);
                if (worksheetPart == null)
                {
                    // Insert a new worksheet of specified name .
                    worksheetPart = InsertWorksheet(xltmDoc.WorkbookPart, sheetName);
                }
                #endregion  // Find desired worksheet by name or create one if it does not exist

                int index;
                Cell cell;
                string cellValue = null;

                #region Loop thru rc_data to write data to excel file

                #region // Insert Header in row 0 //
                for (int r = 0; r < 1; r++)
                {
                    for (int c = 0; c < hdrList.Count(); c++)
                    {
                        #region Insert Values in Cells
                        // Insert the text into the SharedStringTablePart.
                        index = InsertSharedStringItem(hdrList[c], shareStringPart);

                        // Insert a cell into the worksheet.
                        cell = InsertCellInWorksheet(GetExcelColumnName(c + startCol), Convert.ToUInt32(r + startRow), worksheetPart);

                        // Set the value of cell.
                        cell.CellValue = new CellValue(index.ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        cellValue += "\n" + "column: " + GetExcelColumnName(c + startCol) + " : row : " + " " + Convert.ToUInt32(r + startRow) + " : " + index.ToString();

                        #endregion Insert Values in Cells
                    }
                }
                #endregion

                #region //Insert rc_data starting in row 1 //
                for (int r = 1; r < rc_data.Count; r++)
                {
                    for (int c = 0; c < rc_data[r].Count(); c++)
                    {
                        #region Insert Values in Cells
                        // Insert the text into the SharedStringTablePart.
                        index = InsertSharedStringItem(rc_data[r][c], shareStringPart);


                        // Insert a cell into the worksheet.
                        cell = InsertCellInWorksheet(GetExcelColumnName(c + startCol), Convert.ToUInt32(r + startRow), worksheetPart);

                        // Set the value of cell.
                        cell.CellValue = new CellValue(index.ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        cellValue += "\n" + "column: " + GetExcelColumnName(c + startCol) + " : row : " + " " + Convert.ToUInt32(r + startRow) + " : " + index.ToString();

                        #endregion Insert Values in Cells
                    }
                }
                #endregion Loop thru list of Lists

                #endregion

                // Save the new worksheet.
                worksheetPart.Worksheet.Save();
            }  // END using xltmDoc
        }
        #endregion //  write to excel

        #region // InsertSharedStringItem Function to insert text into SharedStringTablePart //
        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }
            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }
        #endregion

        #region // InsertWorksheet Function to add new workseet to excel workbook //
        // Given a WorkbookPart, inserts a new worksheet.
        private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart, string sheetName)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }
        #endregion

        #region // InsertCellinWorksheet Function to insert a cell into worksheet //
        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }
        #endregion

        #region  // Get Data from the KDS Estimation CSV File For Pipes.
        public List<est_data_class> get_KDSEstData_Pipe(string csvFilePath_const)
        {
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
                //TaskDialog.Show("csv_DB", "Pipes : "  + "\n  - Formatted Header: \n\n " + est_db_hdr.ToFormatString() + "\n\n - CSV Format: " + est_db_hdr.ToCsvString_wfname());

                // Get The Rest of the data from FIle and put in a List of est_data_class data-type
                est_data_class_lst = System.IO.File.ReadAllLines(csvFilePath_const)
                                                        .Skip(1)
                                                        .Select(v => est_data_class.FromCsv(v))
                                                        .ToList();
                //TaskDialog.Show("csv_DB", "Pipes : " + "\n - Formatted Data:  \n\n" + est_data_class_lst[2].ToFormatString() + "\n\n - CSV Format: " + est_data_class_lst[2].ToCsvString_wfname());
                //!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! The Contents seems to be the same regardless of BIC   !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            }
            catch (Exception ex)
            {
                TaskDialog.Show("csv_DB", " Exception in accessing KDS Estimation File:" + ex);
            }
            // This is for DEBUG Only.... comment out later.
            string tds_edc = "Pipe System: " + csvFilePath_const + ":: CSV File toList Count: " + est_data_class_lst.Count;
            int edc_cnt = 0;
            foreach (est_data_class est_data_class in est_data_class_lst)
            {
                edc_cnt++;
                tds_edc += "\n " + edc_cnt + "- Formatted Data: " + est_data_class.ToCsvString_wfname();
            }
           // TaskDialog.Show("Csv_to_DB", tds_edc + "\n\n\n\n\n\n\n\n\n");

            return est_data_class_lst;
        }// End of get_KDSEstData_Pipe
        #endregion  // Get Data from the KDS Estimation CSV File.


        #region // GetWorksheetPartByName Function //
        private static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument xltmDoc, string sheetName)
        {
            IEnumerable<Sheet> sheets =
               xltmDoc.WorkbookPart.Workbook.GetFirstChild<Sheets>().
               Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets?.Count() == 0)
            {
                // The specified worksheet does not exist.

                return null;
            }

            string relationshipId = sheets?.First().Id.Value;

            WorksheetPart worksheetPart = (WorksheetPart)xltmDoc.WorkbookPart.GetPartById(relationshipId);

            return worksheetPart;
        }
        #endregion

        #region // GetExcelColumnName Function to convert column number to letter(s) //
        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }
        #endregion

        #region // AppendFileNumberIfExists Function to Append File name //
        // <summary>
        /// A function to add an incremented number at the end of a file name if a file already exists. 
        /// </summary>
        /// <param name="file">The file. This should be the complete path.</param>
        /// <param name="ext">This can be empty.</param>
        /// <returns>An incremented file name. </returns>
        private static string AppendFileNumberIfExists(string file, string ext)
        {
            // This had a VB tidbit that helped to get this started. 
            // http://www.codeproject.com/Questions/212217/increment-filename-if-file-exists-using-csharp

            // If the file exists, then do stuff. Otherwise, we just return the original file name.
            if (File.Exists(file))
            {
                string folderPath = Path.GetDirectoryName(file); // The path to the file. No sense in dealing with this unecessarily. 
                string fileName = Path.GetFileNameWithoutExtension(file); // The file name with no extension. 
                string extension = string.Empty; // The file extension. 
                                                 // This lets us pass in an empty string for the file extension if required. i.e. It just makes this function a bit more versatile. 
                if (ext == string.Empty)
                {
                    extension = Path.GetExtension(file);
                }
                else
                {
                    extension = ext;
                }

                // at this point, find out if the fileName ends in a number, then get that number.
                int fileNumber = 0; // This stores the number as a number for us. 
                                    // need a regex here - \(([0-9]+)\)$
                Regex r = new Regex(@"\(([0-9]+)\)$"); // This matches the pattern we are using, i.e. ~(#).ext
                Match m = r.Match(fileName); // We pass in the file name with no extension.
                string addSpace = " "; // We'll add a space when we don't have our pattern in order to pad the pattern.
                if (m.Success)
                {
                    addSpace = string.Empty; // We have the pattern, so we don't add a space - it has already been added. 
                    string s = m.Groups[1].Captures[0].Value; // This is the single capture that we are looking for. Stored as a string.
                                                              // set fileNumber to the new number.
                    fileNumber = int.Parse(s); // Convert the number to an int.
                                               // remove the numbering from the string as we're constructing it again below.
                    fileName = fileName.Replace("(" + s + ")", "");
                }

                // Start looping. 
                do
                {
                    fileNumber += 1; // Increment the file number that we have above. 
                    file = Path.Combine(folderPath, // Combine it all.
                                            string.Format("{0}{3}({1}){2}", // The pattern to combine.
                                                                      fileName,         // The file name with no extension. 
                                                                      fileNumber,       // The file number.
                                                                      extension,        // The file extension.
                                                                      addSpace));       // A space if needed to pad the initial ~(#).ext pattern.
                }
                while (File.Exists(file)); // As long as the file name exists, keep looping. 
            }
            return file;
        }   // End Of AppendFileNumberIfItExists
        #endregion   // End Of AppendFileNumberIfExists Function to Append File name //


        #region   // I thought i will use this to put my pipe information as sharedParameters, but i think it is the wrong approach.... NOT USED
        // I left it here in case i needed this in the future, i.e. creating Shared Parameterd from within Revit.
        // Taken from the web... needs cleanup
        /// <summary>
        /// Create a new shared parameter
        /// </summary>
        /// <param name="doc">Document</param>
        /// <param name="cat">Category to bind the parameter definition</param>
        /// <param name="nameSuffix">Parameter name suffix</param>
        /// <param name="typeParameter">Create a type parameter? If not, it is an instance parameter.</param>
        /// <returns></returns>
        bool CreateSharedParameter(
            Autodesk.Revit.DB.Document doc,
            Autodesk.Revit.DB.Category cat,
            int nameSuffix,
            ParameterType _deftype,
            bool typeParameter)
        {

            Autodesk.Revit.ApplicationServices.Application app = doc.Application;

            Autodesk.Revit.Creation.Application ca = app.Create;

            const string sharedParameter_dir = "Z:\\BIM\\KDS_TEMPLATE\\2022\\";
            const string sharedParameter_fn = sharedParameter_dir + "KDS_SHARED_PARAMS.txt";
            string _filename = sharedParameter_dir + sharedParameter_fn;

            // get or set the current shared params filename:

            string filename = app.SharedParametersFilename;
            string _defname = "KDS_TEST_";
            string _groupname = "KDS_GROUP_";

            if (0 == filename.Length)
            {
                string path = _filename;
                StreamWriter stream;
                stream = new StreamWriter(path);
                stream.Close();
                app.SharedParametersFilename = path;
                filename = app.SharedParametersFilename;
            }

            // get the current shared params file object:

            DefinitionFile file
          = app.OpenSharedParameterFile();

            if (null == file)
            {
                TaskDialog.Show("CreateSharedParameter", "Error getting the shared params file.");

                return false;
            }

            // get or create the shared params group:

            DefinitionGroup group
          = file.Groups.get_Item(_groupname);

            if (null == group)
            {
                group = file.Groups.Create(_groupname);
            }

            if (null == group)
            {
                TaskDialog.Show("CreateSharedParameter", "Error getting the shared params group.");

                return false;
            }

            // set visibility of the new parameter:

            // Category.AllowsBoundParameters property
            // indicates if a category can have user-visible
            // shared or project parameters. If it is false,
            // it may not be bound to visible shared params
            // using the BindingMap. Please note that
            // non-user-visible parameters can still be
            // bound to these categories.

            ////////bool visible = cat.AllowsBoundParameters;

            // get or create the shared params definition:

            string defname = _defname + nameSuffix.ToString();


            Definition definition = group.Definitions.get_Item(
          defname);

            if (null == definition)
            {
                //definition = group.Definitions.Create( defname, _deftype, visible ); // 2014

                ExternalDefinitionCreationOptions opt
                  = new ExternalDefinitionCreationOptions(defname, _deftype);

                opt.Visible = true;  // visible;

                definition = group.Definitions.Create(opt); // 2015
            }
            if (null == definition)
            {
                TaskDialog.Show("CreateSharedParameter", "Error creating shared parameter.");

                return false;
            }

            /*// create the category set containing our category for binding:

            CategorySet catSet = ca.NewCategorySet();
            catSet.Insert(cat);

            // bind the param:

            try
            {
                Autodesk.Revit.DB.Binding binding = typeParameter
                  ? ca.NewTypeBinding(catSet) as Autodesk.Revit.DB.Binding
                  : ca.NewInstanceBinding(catSet) as Autodesk.Revit.DB.Binding;

                // we could check if it is already bound,
                // but it looks like insert will just ignore
                // it in that case:

                doc.ParameterBindings.Insert(definition, binding);

                // we can also specify the parameter group here:

                //doc.ParameterBindings.Insert( definition, binding,
                //  BuiltInParameterGroup.PG_GEOMETRY );

                Debug.Print(
                  "Created a shared {0} parameter '{1}' for the {2} category.",
                  (typeParameter ? "type" : "instance"),
                  defname, cat.Name);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("CreateSharedParameter", string.Format("Error binding shared parameter to category {0}: {1}",
                  cat.Name, ex.Message));
                return false;
            }*/
            return true;
        }  // End Of CreateSharedParameter


        #endregion   // Endo Of CreateSharedParameter






    }   // End of Class Export_To_Excel 
}   // End Of Namespace 
