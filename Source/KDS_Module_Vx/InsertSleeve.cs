using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Windows.Forms;




namespace KDS_Module_Vx
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class InsertSleeve : IExternalCommand
    {
        string _sleeveString = "";
        List<string> _sleeveStr_lst = new List<string> {
        "KDS_Hilti-FS_CFS_CID",
"KDS_Husky-Band_SD4000",
"KDS_Hilti-FS_CFS_CID_AsFixture",
"KDS_Hilti-FS_CFS_CID_AsMechCouplFitt",
"KDS_Hilti-FS_CFS_CID_AsFittUnion"
    };




        Dictionary<string, double> flrTh_dict = new Dictionary<string, double>()
                {
            {"0.5", 0.5},{"1/2", 0.5},
            {"0.75", 0.75},{"3/4", 0.75},
            {"1.0", 1.0},{"1", 1.0},
            {"1.25", 1.25},{"1 1/4", 1.25},
            {"1.5", 1.5},{"1 1/2", 1.5},
            {"1.75", 1.75},{"1 3/4", 1.75},
            {"2.0", 2.0},{"2", 2.0},
            {"2.25", 2.25},{"2 1/4", 2.25},
            {"2.5", 2.5},{"2 1/2", 2.5},
            {"2.75", 2.75},{"2 3/4", 2.75},
            {"3.0", 3.0},{"3", 3.0},
            {"3.25", 3.25},{"3 1/4", 3.25},
            {"3.5", 3.5},{"3 1/2", 3.5},
            {"3.75", 3.75},{"3 3/4", 3.75},
            {"4.0", 4.0},{"4", 4.0},
            {"4.25", 4.25},{"4 1/4", 4.25},
            {"4.5", 4.5},{"4 1/2", 4.5},
            {"4.75", 4.75},{"4 3/4", 4.75},
            {"5.0", 5.0},{"5", 5.0},
            {"5.5", 5.5},
            {"6.0", 6.0},{"6", 6.0},
            {"6.5", 6.5},
            {"7.0", 7.0},{"7", 7.0},
            {"7.5", 7.5},
            {"8.0", 8.0},{"8", 8.0},
            {"8.5", 8.5},
            {"9.0", 9.0},{"9", 9.0},
            {"9.5", 9.5},
            {"10.0", 10.0},{"10", 10.0},
            {"10.5", 10.5},
            {"11.0", 11.0},{"11", 11.0},
            {"11.5", 11.5},
            {"12.0", 12.0},{"12", 12.0},
            {"12.5", 12.5}
                };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            #region // Some Definitions 
            Autodesk.Revit.DB.Document actvDoc = commandData.Application.ActiveUIDocument.Document;
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            UIApplication uiApp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiApp.Application;
            List<Autodesk.Revit.DB.Point> sleevesLocPnt_lst = new List<Autodesk.Revit.DB.Point>();
            #endregion  // Some Definitions
           
            
            #region  // Test guid 
           /* Guid KDS_PIPING_SYSTEM_TYPE_GUID = "d7196377 - 8d4d - 4b25 - 8789 - 912f72cc2f80";
            List<Element> sharedParams = new List<Element>();
            FilteredElementCollector collector
                = new FilteredElementCollector(actvDoc)
                .WhereElementIsNotElementType();
            // Filter elements for shared parameters only
            collector.OfClass(typeof(SharedParameterElement));
            string guid_str = "GUID Parameter Values \n";

            foreach (Element e in collector)
            {
                SharedParameterElement param_sharedElem = e as SharedParameterElement;
                Definition def = param_sharedElem.GetDefinition();
                string param_str = GetParamValueByGuid( param_sharedElem.GuidValue, param_sharedElem)??"NA-guid";

                //Debug.WriteLine("[" + e.Id + "]\t" + def.Name + "\t(" + param.GuidValue + ")");
                guid_str += "id= [" + e.Id + "]\n" +
                             "Name: " + def.Name + "\n" +
                             "GUIDValue: (" + param_sharedElem.GuidValue + ")" + "\n" +
                             "parm Value: " + param_str + "\n" +
                             "------------------------------\n";
            }
            TaskDialog.Show("insertSleeve",guid_str);
*/
            #endregion

            #region  // Get Sleeve Family Name to Use from the User  -- now hardcoded to "KDS_Hilti-FS_CFS_CID"
            /*string sleeveInputStr = "Please Select a Family Name:\n";
            int sleeveInputStr_i = 0;
            int i = 0;
            foreach (string st in _sleeveStr_lst) { i++; sleeveInputStr += i + "- " + st + "\n"; }

            if (int.TryParse(Interaction.InputBox(sleeveInputStr, "InsertSleeve", "Select a Family", 500, 500), out sleeveInputStr_i))
            {
                _sleeveString = _sleeveStr_lst[sleeveInputStr_i - 1];
            }
            else return Result.Succeeded;*/


            _sleeveString = _sleeveStr_lst[0];

            #endregion  // Get Sleeve Family Name to Use from the User


            #region   // Get Family of Sleeves and List of all Pipes in Host Doc.sleevesLocPnt_lst
            // Get The FamilySymbol for the Sleeve
            FilteredElementCollector sleeveCollector = new FilteredElementCollector(actvDoc);
            Family sleeveFam = sleeveCollector.OfClass(typeof(Family)).OfType<Family>().FirstOrDefault(f => f.Name.Equals(_sleeveString));

            FamilySymbol sleeveFamSymbol = actvDoc.GetElement(sleeveFam.GetFamilySymbolIds().FirstOrDefault()) as FamilySymbol;

            string flrTh_str = null;
            string insTh_str = null;

            // Get a List of All Pipes 
            FilteredElementCollector pipeCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(Pipe));
            List<Pipe> pipes_lst = pipeCollector.Cast<Pipe>().ToList();

            #region // Debub only  -- ommited
            /*string inThStr = "Pipe Diam || Ins_dbl || overallSize_dbl ||\n";
            foreach (Pipe pipe in pipes_lst)
            {
                inThStr += pipe.Diameter * 12.0 + "||" +
                     get_PipeInsulationThickness(pipe) * 12.0 + "||" +
                     get_PipeOverallSize(pipe) * 12.0 + "||" +
                    "\n";
            }
            TaskDialog.Show("insertSleeve", inThStr);*/
            #endregion // Debug only 

            //Get all Pipes that are of an overall diameter + insulation Size within 2" and 6" size
            // Here we will get all pipes with overall size less than (0.5') 6" since after you add the 1" gap to the sleeve, you get 8",
            // and they do not make sleeves larger than 8".
            // Talking to draino, he says just put 2" Sleeves anyways, for pipes with over all size Less than 2".
            // So This will be added during the creation of the Sleeves. anything less than 2" overall size will get a 2" Sleeve.
            //List<Pipe> pipes_OAsize_dbl_lst = pipes_lst.Where(p => get_PipeOverallSize(p) > 0.146 && get_PipeOverallSize(p) < 0.5).ToList();
            List<Pipe> pipes_OAsize_dbl_lst = pipes_lst.Where(p => get_PipeOverallSize(p) < 0.51).ToList();

            #region // Debug Only -- ommited
            /*            string inThStr0 = "Pipe Diam || Ins_dbl || overallSize_dbl ||\n";

                        foreach (Pipe pipe in pipes_OAsize_dbl_lst)
                        {
                            inThStr0 += pipe.Diameter * 12.0 + "||" +
                                 get_PipeInsulationThickness(pipe) * 12.0 + "||" +
                                 get_PipeOverallSize(pipe) * 12.0 + "||" +
                                "\n";
                        }
                        TaskDialog.Show("insertSleeve", inThStr0);*/
            #endregion // End Of Debug Only -- ommited


            // Get pipes that are Vertical
            List<Pipe> pipes_size_slope_lst = pipes_OAsize_dbl_lst.Where(p => IsVertical(p)).ToList<Pipe>();


            #region // Debug Only -- ommited
            /*string inThStr2 = "Pipe Diam || Ins_dbl || overallSize_dbl || ElementID ||\n";

            foreach (Pipe pipe in pipes_size_slope_lst)
            {
                inThStr2 += pipe.Diameter * 12.0 + "||" +
                     get_PipeInsulationThickness(pipe) * 12.0 + "||" +
                     get_PipeOverallSize(pipe) * 12.0 + "||" +
                     pipe.Id + "||" +
                    "\n";
            }
            TaskDialog.Show("insertSleeve", inThStr2);*/
            #endregion // End Of Debug Only -- ommited


            #endregion   // Get Family of Sleeves and List of all Pipes in Host Doc.



            // For Floors we want to work with the floors from the Architectural model... We will copy and paste these floors in the host DOc sine this is the only way to get their ids in another step.
            // if this is not available, then we will create temp Floors  insert Sleeves and pipe-floor instersections, then delete these floors.


            #region // Logic Statements to determine method of inserting sleeves //
            // //Check if the Host Doc has any floors.
            FilteredElementCollector hostDoc_floors_col = new FilteredElementCollector(actvDoc).OfClass(typeof(Floor));  //FLOOR_PARAM_IS_STRUCTURAL
            List<Floor> hostDoc_floors_lst = hostDoc_floors_col.Cast<Floor>().ToList();
            List<Floor> lnkdDocFlr_lst = new List<Floor>();

            if (hostDoc_floors_lst.Count != 0 && hostDoc_floors_lst != null)    // --- Host Doc has Floors... So no need to create or Load from linked Docs
            {
                TaskDialog.Show("insertSleeves", " Floors Found in Document. \n No Need to Create any Temp Floors. \n Host Document floors Count = " + hostDoc_floors_lst.Count);
                insertSleeves(commandData.Application, pipes_size_slope_lst, hostDoc_floors_lst, sleeveFamSymbol, insTh_str);

            }
            else   // No Floors in Host Doc, so Check the Linked Models
            {


                DialogResult dialogResult = MessageBox.Show("Load Linked Documents: \n\n Yes: Loads All Linked Documnets. \n No: Give you an option to Select what to load. \n Cancel: Does Not Load Any Documents.", "Load Linked Documents", MessageBoxButtons.YesNoCancel);
                if (dialogResult == DialogResult.Yes || dialogResult == DialogResult.No)
                {
                    List<RevitLinkType> loadedDocsLinkTypes_lst = new List<RevitLinkType>();

                    if (dialogResult == DialogResult.Yes)
                    {

                        loadedDocsLinkTypes_lst = load_all_Un_LnkdDocs(actvDoc);
                    }
                    else { loadedDocsLinkTypes_lst = load_Selected_Un_LnkdDocs(actvDoc); }

                    #region   // Loop to dispaly All loaded Linked Document.
                    string ldlnkd_str = "";
                    int ldlnkd_cnt = 0;
                    ldlnkd_str = "--- loadedDocsLinkTypes_lst.Count: " + loadedDocsLinkTypes_lst.Count + " ---";
                    foreach (RevitLinkType rvtlnk in loadedDocsLinkTypes_lst)
                    {
                        ldlnkd_cnt++;
                        ldlnkd_str += "\n " + ldlnkd_cnt + "- rvtlnk.id: " + rvtlnk.Id + " - rvtlnk.Name: " + rvtlnk.Name;
                    }
                    TaskDialog.Show("insertSleeves", ldlnkd_str);
                    #endregion   // End of Loop to dispaly All loaded Linked Document.

                    List<flrsPerLnkdDoc_strct> flrs_kvp_lst = getFlrs_AllLnkdDocs_lst(app, BuiltInCategory.OST_Floors);

                    #region   // Loop to dispaly All Floors within each Linked Document.
                    string tds = "";
                    int dc = 0;

                    int tf = 0;

                    foreach (flrsPerLnkdDoc_strct flrs_kvp in flrs_kvp_lst)
                    {
                        dc++;
                        int flr_cnt = 0;
                        tds += " \n " + dc + "- Floors in : " + flrs_kvp.lnkDoc.Title;
                        foreach (ElementId flr_eid in flrs_kvp.flrsElId_lst)
                        {
                            flr_cnt++;
                            tf++;
                            Floor flr = flrs_kvp.lnkDoc.GetElement(flr_eid) as Floor;
                            if (flr == null) { }
                            else
                            {
                                tds += "\n   " + flr_cnt + "- Floor Name: " + flr.Name;
                            }
                        }
                    }
                    tds = "--- List of All Floor Names per Linked Document ---" + "\n         Total Docs: " + dc + " Total Floors: " + tf + tds;
                    TaskDialog.Show("insertSleeves", tds);
                    #endregion   // End of  Loop to dispaly All Floors within each Linked Document.

                    #region // Copy all Floors Found in Linked Documents to host Document (actvDoc) .
                    foreach (flrsPerLnkdDoc_strct flrs_kvp in flrs_kvp_lst)
                    {
                        TaskDialog.Show("insertSleeves", "- Copying " + flrs_kvp.flrsElId_lst.Count + " Floors from Linked Doc: " + flrs_kvp.lnkDoc.Title);


                        //resetFloor(actvDoc, flrs_kvp.flrsElId_lst); 
                        copyPasteIds(actvDoc, flrs_kvp.lnkDoc, flrs_kvp.flrsElId_lst);

                    }
                    #endregion // End of Copy all Floors Found in Lined Documents to host Document (actvDoc) .

                    //return Result.Succeeded;

                    unload_loadedDocs_lst(loadedDocsLinkTypes_lst);

                    FilteredElementCollector lnkdDocFlr_col = new FilteredElementCollector(actvDoc).OfClass(typeof(Floor));  //FLOOR_PARAM_IS_STRUCTURAL
                    lnkdDocFlr_lst = hostDoc_floors_col.Cast<Floor>().ToList();
                }

                else if (dialogResult == DialogResult.No)
                {

                    //TaskDialog.Show("insertSleeves", "Do Not Load Linked Documents: " + DialogResult.No);
                }

                if (lnkdDocFlr_lst.Count != 0 && lnkdDocFlr_lst != null)    // Linked Models has Floors so use them
                                                                            //if (lnkdDocFlr_col.Count != 0 && lnkdDocFlr_col != null)    // Linked Models has Floors so use them
                {
                    FilteredElementCollector hostDoc_cpdFlrs_col = new FilteredElementCollector(actvDoc).OfClass(typeof(Floor));  //FLOOR_PARAM_IS_STRUCTURAL
                    List<Floor> hostDoc_cpdFlrs_lst = hostDoc_floors_col.Cast<Floor>().ToList();

                    resetFloor(actvDoc, hostDoc_cpdFlrs_lst);

                    TaskDialog.Show("insertSleeves", " Found " + hostDoc_cpdFlrs_lst.Count + " Floors in Host Document, after copying Floors from Linked Documents. \n No Need to Create any Temp Floors.");

                    insertSleeves(commandData.Application, pipes_size_slope_lst, hostDoc_cpdFlrs_lst, sleeveFamSymbol, insTh_str);


                    TaskDialog.Show("insertSleeves", " Deleting All Copyed Floors from Linked Documents.");
                    DeleteFloors(actvDoc, hostDoc_cpdFlrs_lst);
                }

                else   // Neither Host Doc, Nor Linked Docs have Floors, so create ones based on host Doc Levels and get instersections.
                {

                    flrTh_str = Interaction.InputBox("Prompt", "Floor Thickness", "6.5", 800, 800);
                    //insTh_str = Interaction.InputBox("Prompt", "Insulation Thickness", "1.5", 800, 800);



                    TaskDialog.Show("insertSleeves", "    Neither Host Doc, Nor Linked Docs have Floors defined.  \n - Create Temp Floors based on host Doc Levels. \n - Place Sleeves at instersections with pipes. \n - Delete Temp Floors");
                    List<Floor> levelsFloors_lst = new List<Floor>();
                    // Create Temp Floors in Host DOc
                    levelsFloors_lst = insertLevelFloors(commandData.Application, actvDoc, flrTh_str);
                    TaskDialog.Show("insertSleeves", "Created new Temp Floors. \n levelsFloors_lst.Count = " + levelsFloors_lst.Count);
                    // Insert Floor Sleeves ata intersections
                    insertSleeves(commandData.Application, pipes_size_slope_lst, levelsFloors_lst, sleeveFamSymbol, insTh_str);

                    TaskDialog.Show("insertSleeves", " Delete Temp Floors");
                    // Delete Temp Floors in Host DOc
                    DeleteFloors(actvDoc);
                }


            }

            #endregion

            return Result.Succeeded;
        }



        #region // insertSleeves Function to place sleeve where Pipe and Level intersect --- No Floor exists here. so we need to create floors as well//
        public List<XYZ> insertSleeves(UIApplication uiApp, List<Pipe> pipes_lst, List<Floor> floors_lst, FamilySymbol sleeveFamSymbol, string insTh_str)
        {
            #region // Defining and Collection of Elements and Variables //

            UIDocument uidoc = uiApp.ActiveUIDocument;
            Autodesk.Revit.DB.Document actvDoc = uidoc.Document;


            List<XYZ> points = new List<XYZ>();
            List<XYZ> sleevesLocPnt_lst = new List<XYZ>();
            //return sleevesLocPnt_lst; /// Debug
            #endregion


            #region // Insert Sleeve where Pipe and floor intersect //

            var watch = System.Diagnostics.Stopwatch.StartNew();
            int index = 0;

            //FilteredElementCollector floorCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(Floor));
            //List<Floor> floors = floorCollector.Cast<Floor>().ToList();

            /*            TaskDialog.Show("insertSleeves", "In insertSleeves Function " +
                                                         "\nfloors_lst.Count: " + floors_lst.Count +
                                                         "\npipes_lst.Count: " + pipes_lst.Count);*/
            FamilyInstance sleeve_famInst = null;

            foreach (Pipe p in pipes_lst)
            {
                foreach (Floor f in floors_lst)
                {
                    Curve pipeCurve = FindPipeCurve(p);

                    XYZ intersection = null;
                    XYZ xyz_tol = new XYZ(0.01, 0.01, 1);

                    List<Face> floorFaces = FindFloorFace(f);
                    Face frstFlrFace = floorFaces.FirstOrDefault();


                    //for (int i = 0; i < 1; i++)
                    if (frstFlrFace != null)
                    {
                        //intersection = FindFaceCurve(pipeCurve, floorFaces[i]);
                        intersection = FindFaceCurve(pipeCurve, frstFlrFace);
                        if (null == intersection)
                        {
                            //TaskDialog.Show("insertSleeves", "  In insertSleeves Function " + "\nintersection is NULL !");
                        }

                        //else if (sleevesLocPnt_lst.Contains(intersection))
                        else if (isWithinRange(sleevesLocPnt_lst, intersection, xyz_tol))
                        {
                            //TaskDialog.Show("insertSleeves", "  In insertSleeves Function " + "\n intersection is Found in List, So no need to insert Element again. !");
                        }
                        else //(null != intersection)
                        {
                            //TaskDialog.Show("insertSleeves", "  In insertSleeves Function " + "\nInside the !=null intersection block !");
                            index++;
                            using (Transaction insertSleeve_trx = new Transaction(actvDoc, "Insert Sleeves"))
                            {
                                insertSleeve_trx.Start();
                                sleeveFamSymbol.Activate();
                                sleeve_famInst = actvDoc.Create.NewFamilyInstance(intersection, sleeveFamSymbol, f, StructuralType.NonStructural);
                                insertSleeve_trx.Commit();
                                sleevesLocPnt_lst.Add(intersection);
                            }

                            using (Transaction setSleeveParam_trx = new Transaction(actvDoc, "Set Sleeve Parameters - Diameter and System Type"))
                            {
                                setSleeveParam_trx.Start();
                                try
                                {
                                    // Set failure handler  .. this is to handle me writing the same value in the Mark parameter which i guess, is not expected.
                                    var failureOptions = setSleeveParam_trx.GetFailureHandlingOptions();
                                    failureOptions.SetFailuresPreprocessor(new FailurePreproccessor());
                                    setSleeveParam_trx.SetFailureHandlingOptions(failureOptions);

                                    //TaskDialog.Show("insertSleeve", " Just Disabled Warnings on Duplicate Mark Values");

                                    //Parameter OverallSize_p = p.get_Parameter(BuiltInParameter.RBS_REFERENCE_OVERALLSIZE);  //  Gets the sum of the Pipe Diameter and the Insulation Thickness if any
                                    double OverallSize_dbl = get_PipeOverallSize(p);
                                    //TaskDialog.Show("insertSleeve", " OverallSize: " + OverallSize_dbl);

                                    double slv_diam_dbl_adj = 0.0;

                                    if (OverallSize_dbl < 0.158) { slv_diam_dbl_adj = 0.167; }          // For  than 1.9" => 2"   
                                    if (OverallSize_dbl > 0.158 && OverallSize_dbl < 0.200) { slv_diam_dbl_adj = 0.334; }    // For Between 1.9" and 2.4" => 2" => 4" Hilti0.334
                                    if (OverallSize_dbl > 0.200 && OverallSize_dbl < 0.283) { slv_diam_dbl_adj = 0.417; }   // For Between 2.4" and 3.4" => 3" => 5" Hilti0.417
                                    if (OverallSize_dbl > 0.283 && OverallSize_dbl < 0.367) { slv_diam_dbl_adj = 0.500; }   // For Between 3.4" and 4.4" => 4" => 6" Hilti0.500
                                    if (OverallSize_dbl > 0.367 && OverallSize_dbl < 0.508) { slv_diam_dbl_adj = 0.667; }   // For Between 4.4" and 6.1" => 6" => 8" Hilti0.667
                                    if (OverallSize_dbl > 0.508) { slv_diam_dbl_adj = 0.667; TaskDialog.Show("insertSleeve", "Hilti Does Not Support Sleeves Larger Than 6 inch"); }   // For Between 2" and 3" => 3"




                                    //TaskDialog.Show("insertSleeve", " slv_diam_dbl_adj: " + slv_diam_dbl_adj);

                                    // Get the Diameter Parameter for the Sleeve THEN set it to the size from the relevant pipe.
                                    Parameter slv_diam = sleeve_famInst.LookupParameter("Size");

                                    // Set the Diameter for the Sleeve
                                    slv_diam.Set(slv_diam_dbl_adj);
                                    /*if (slv_diam.Set(slv_diam_dbl_adj) == true)
                                        TaskDialog.Show("insertSleeve", " Set Param slv_diam to: " + slv_diam.AsValueString());
                                    else
                                        TaskDialog.Show("insertSleeve", " Could NOT set Param Value slv_diam to : " + slv_diam.AsValueString());
*/


                                    // Get the System Type  Parameter for the Sleeve THEN set it to the that of the relevant pipe.
                                    Parameter p_SysType = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);

                                    //TaskDialog.Show("insertSleeve", " Pipe p_SysType: " + p_SysType.AsValueString());

                                    // Get the parameter of the Sleeve - PipeAccessory 
                                    //Parameter slv_SysType = sleeve_famInst.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                                    Parameter slv_SysType = sleeve_famInst.LookupParameter("KDS_PIPING_SYSTEM_TYPE");

                                    Guid slv_SysType_GUID = slv_SysType.GUID;
                                    Parameter p3 = sleeve_famInst.get_Parameter(slv_SysType_GUID);



                                    TaskDialog.Show("insrtSleeve", "GUID for  KDS_PIPING_SYSTEM_TYPE: " + slv_SysType.GUID);
                                    TaskDialog.Show("insrtSleeve", "p3 value: " + p3.AsString());

                                    #region // Debug Only Check if readOnly and UserModifiable. if it is , then api cannot change it, only UI can.
                                    if (slv_SysType.IsReadOnly) TaskDialog.Show("insrtSleeve", "Sleeve KDS_PIPING_SYSTEM_TYPE: Is Read Only ");
                                    else TaskDialog.Show("insrtSleeve", "Sleeve KDS_PIPING_SYSTEM_TYPE: Is NOT  Read Only ");

                                    if (slv_SysType.UserModifiable) TaskDialog.Show("insrtSleeve", "Sleeve KDS_PIPING_SYSTEM_TYPE: Is UserModifiable ");
                                    else TaskDialog.Show("insrtSleeve", "Sleeve KDS_PIPING_SYSTEM_TYPE: Is NOT  UserModifiable");

                                    #endregion  // End Of Debug Only Check if readOnly and UserModifiable. if it is , then api cannot change it, only UI can.
                                    TaskDialog.Show("insrtSleeve", "BEFORE set Sleeve KDS_PIPING_SYSTEM_TYPE: " + slv_SysType.AsValueString());
                                    TaskDialog.Show("insrtSleeve", "BEFORE set Sleeve KDS_PIPING_SYSTEM_TYPE w GUID: " + actvDoc.GetElement(sleeve_famInst.Id).get_Parameter(slv_SysType_GUID).AsValueString());
                                    // Set Sleeve RBS_PIPING_SYSTEM_TYPE_PARAM Value
                                    slv_SysType.Set(p_SysType.AsValueString());

                                    //TaskDialog.Show("insrtSleeve", "AFTER set RBS_PIPING_SYSTEM_TYPE_PARAM: " + sleeve_famInst.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString());
                                    TaskDialog.Show("insrtSleeve", "AFTER set KDS_PIPING_SYSTEM_TYPE: " + sleeve_famInst.LookupParameter("KDS_PIPING_SYSTEM_TYPE").AsValueString());
                                    TaskDialog.Show("insrtSleeve", "AFTER set Sleeve KDS_PIPING_SYSTEM_TYPE w GUID: " + actvDoc.GetElement(sleeve_famInst.Id).get_Parameter(slv_SysType_GUID).AsValueString());


                                  p3 = sleeve_famInst.get_Parameter(slv_SysType_GUID);



                                    TaskDialog.Show("insrtSleeve", "GUID for  KDS_PIPING_SYSTEM_TYPE: " + slv_SysType.GUID);
                                    TaskDialog.Show("insrtSleeve", "p3 value: " + p3.AsString());

/*                                    SharedParameterElement shParamElement = SharedParameterElement.Lookup(actvDoc, new Guid("d7196377-8d4d-4b25-8789-912f72cc2f80"));

                                    ParameterValueProvider pvp = new ParameterValueProvider(shParamElement.Id);
                                    TaskDialog.Show("insrtSleeve", "pvp value: " + pvp.Parameter.ToString());
                                    FilterStringRuleEvaluator evaluator = new FilterStringEquals(); 
                                    FilterRule rule = new FilterStringRule(pvp, evaluator, "KDS_PIPING_SYSTEM_TYPE", false);
                                    */
                                    Element element = actvDoc.GetElement(sleeve_famInst.Id);
                                    Parameter p4 = element.get_Parameter(new Guid("d7196377-8d4d-4b25-8789-912f72cc2f80"));
                                    TaskDialog.Show("insrtSleeve", "p4 AsString: " + p4.AsString());
                                    TaskDialog.Show("insrtSleeve", "p4 AsValueString: " + p4.AsValueString());
                                    TaskDialog.Show("insrtSleeve", "p4 ToString: " + p4.ToString());

                                }


                                catch (Exception e)
                                {

                                    TaskDialog.Show("insrtSleeve", "  In insertSleeves Function " + "\nException try to write parameters: e: " + e);
                                }
                                setSleeveParam_trx.Commit();
                                setSleeveParam_trx.Dispose();
                                //TaskDialog.Show("insrtSleeve", "AFTER set RBS_PIPING_SYSTEM_TYPE_PARAM: " + sleeve_famInst.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString());
                                //TaskDialog.Show("insrtSleeve", "AFTER COMMIT set KDS_PIPING_SYSTEM_TYPE: " + sleeve_famInst.get_Parameter(sleeve_famInst.LookupParameter("KDS_PIPING_SYSTEM_TYPE").GUID).AsValueString());
                               

                            }
                        }
                    }   // if frstflrface
                }
            }
            watch.Stop();
            var elapsedTime = watch.ElapsedMilliseconds;
            TaskDialog.Show("test", "  In insertSleeves Function " + "\n# of sleeves inserted: " + index + "\n execution time : " + elapsedTime / 1000.0 + " seconds");
            #endregion

            return sleevesLocPnt_lst;
        }  // End of insertSleeves
        #endregion  // End of insertSleeves


        public string GetParamValueByGuid(Guid guid, Element e)
        {
            var paramValue = string.Empty;

            foreach (Parameter parameter in e.Parameters)
            {
                if (parameter.IsShared)
                {
                    if (parameter.GUID == guid)
                    {
                        paramValue = parameter.AsString();
                    }
                }
            }

            return paramValue;
        }


        public Parameter GetParameter(Element e, Guid guid)
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




        #region // Reset the shape of the floors as Elementid 
        public void resetFloor(Autodesk.Revit.DB.Document actvDoc, List<ElementId> floors_lst)
        {
            // I Came Across floors that had their shape modified, and the i was not able to insert a sleeve.
            // So i reset their shape and i was able to insert a sleeve.  this loop does that progrmamtically. (select Floor=>Modify=>reset shape)

            foreach (ElementId ef in floors_lst)
            {
                Floor f = actvDoc.GetElement(ef) as Floor;
                //TaskDialog.Show("insertSleeves", " FOUND IT Floor Element Name : " + f.Name);
                using (Transaction ResetSlabShape_trx = new Transaction(actvDoc, "Reset the Shape of the floor"))
                {
                    try
                    {
                        ResetSlabShape_trx.Start();
                        f.SlabShapeEditor.ResetSlabShape();

                    }
                    catch (Exception e)
                    {

                        TaskDialog.Show("insrtSleeve", "  In ResetSlabShape_trx  " + "\nException try to Reset Floor shape: e: " + e);
                    }
                    ResetSlabShape_trx.Commit();
                }
            }
        }
        #endregion  // End of Reset the shape of the floors.



        #region // Reset the shape of the floors as Floor 
        public void resetFloor(Autodesk.Revit.DB.Document actvDoc, List<Floor> floors_lst)
        {
            // I Came Across floors that had their shape modified, and the i was not able to insert a sleeve.
            // So i reset their shape and i was able to insert a sleeve.  this loop does that progrmamtically. (select Floor=>Modify=>reset shape)
            int flr_cnt = 0;
            string flr_nm = "";
            foreach (Floor flr in floors_lst)
            {
                flr_cnt++;
                flr_nm = flr.Name;

                if (flr.SlabShapeEditor != null)
                {
                    //TaskDialog.Show("insertSleeves", flr_cnt +" -Reseting the Slab shape of Floor: \n Name : " + flr.Name + "\n ID: " + flr.Id + "\n IsFoundationSlab: " + flr.FloorType.IsFoundationSlab);
                    using (Transaction ResetSlabShape_trx = new Transaction(actvDoc, "Reset the Shape of the floor"))
                    {
                        try
                        {
                            ResetSlabShape_trx.Start();
                            flr.SlabShapeEditor.ResetSlabShape();

                        }
                        catch (Exception e)
                        {

                            TaskDialog.Show("insrtSleeve", "  In ResetSlabShape_trx  " + "\nException try to Reset Floor shape: e: " + e);
                        }

                        ResetSlabShape_trx.Commit();
                    }
                }
                else
                {
                    TaskDialog.Show("insertSleeves", flr_cnt + " -Could Not Reset the Slab shape of Floor: \n Name : " + flr.Name + "\n ID: " + flr.Id + "\n IsFoundationSlab: " + flr.FloorType.IsFoundationSlab);
                }
            }
        }
        #endregion  // End of Reset the shape of the floors.


        #region // InsertFloor Function to create floor //
        public Floor InsertFloor(UIApplication uiApp, Autodesk.Revit.DB.Level lvl, string flrTh_str)
        {
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Autodesk.Revit.DB.Document actvDoc = uidoc.Document;

            Floor floor = null;
            #region // Transaction to Create Floor //
            using (Transaction createFloor = new Transaction(actvDoc, "Create floor"))
            {
                createFloor.Start();
                FloorType floorType = null;
#if RVT2020
                #region // Get a floor type for floor creation 2020
                floorType = new FilteredElementCollector(actvDoc).OfClass(typeof(FloorType)).First(e => e.Name.Equals("Generic - 12\"")) as FloorType;   //For Revit 2021 and older

                // The normal vector (0,0,1) that must be perpendicular to the profile.
                XYZ normal = XYZ.BasisZ;

                XYZ first = new XYZ(1000, 1000, 0);
                XYZ second = new XYZ(1000, -1000, 0);
                XYZ third = new XYZ(-1000, -1000, 0);
                XYZ fourth = new XYZ(-1000, 1000, 0);

                CurveArray profile = new CurveArray();   // For revit 2021 and older
                profile.Append(Line.CreateBound(first, second));
                profile.Append(Line.CreateBound(second, third));
                profile.Append(Line.CreateBound(third, fourth));
                profile.Append(Line.CreateBound(fourth, first));

                floor = actvDoc.Create.NewFloor(profile, floorType, lvl, true, normal);   // For Revit 2021 and older
                Parameter param = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);// For Revit 2022
                param.Set(5.0);
                #endregion  // Get a floor type for floor creation 2020
#elif RVT2021
                XYZ normal = XYZ.BasisZ;

                #region                // Get a floor type for floor creation 2021
                floorType = new FilteredElementCollector(actvDoc).OfClass(typeof(FloorType)).First(e => e.Name.Equals("Generic - 12\"")) as FloorType;   //For Revit 2021 and older

                // The normal vector (0,0,1) that must be perpendicular to the profile.


                XYZ first = new XYZ(1000, 1000, 0);
                XYZ second = new XYZ(1000, -1000, 0);
                XYZ third = new XYZ(-1000, -1000, 0);
                XYZ fourth = new XYZ(-1000, 1000, 0);

                CurveArray profile = new CurveArray();   // For revit 2021 and older
                profile.Append(Line.CreateBound(first, second));
                profile.Append(Line.CreateBound(second, third));
                profile.Append(Line.CreateBound(third, fourth));
                profile.Append(Line.CreateBound(fourth, first));

                floor = actvDoc.Create.NewFloor(profile, floorType, lvl, true, normal);   // For Revit 2021 and older
                Parameter param = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);// For Revit 2022
                param.Set(5.0);
                #endregion                // Get a floor type for floor creation 2021
#elif RVT2022
                #region               // Get a floor type for floor creation 2022
                ElementId floorTypeId = Floor.GetDefaultFloorType(actvDoc, false);   // For Revit 2022

                // The normal vector (0,0,1) that must be perpendicular to the profile.
                XYZ normal = XYZ.BasisZ;

                XYZ first = new XYZ(1000, 1000, 0);
                XYZ second = new XYZ(1000, -1000, 0);
                XYZ third = new XYZ(-1000, -1000, 0);
                XYZ fourth = new XYZ(-1000, 1000, 0);

                CurveLoop profile = new CurveLoop();   // for Revit 2022
                profile.Append(Line.CreateBound(first, second));
                profile.Append(Line.CreateBound(second, third));
                profile.Append(Line.CreateBound(third, fourth));
                profile.Append(Line.CreateBound(fourth, first));

                floor = Floor.Create(actvDoc, new List<CurveLoop> { profile }, floorTypeId, lvl.Id);    // For Revit 2022
                Parameter param = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);// For Revit 2022
                param.Set(5);// For Revit 2022
                #endregion               // Get a floor type for floor creation 2022
#else
#endif

                Parameter p = floor.get_Parameter(BuiltInParameter.LEVEL_PARAM);
                Parameter p1 = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                Parameter p2 = floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_DEFAULT_THICKNESS_PARAM);

                double flrTh_dbl = flrTh_dict[flrTh_str];
                //TaskDialog.Show("InsertFloor", " Floor Thickness is: " + flrTh_dbl);
                p.Set(lvl.Id);
                p1.Set(0);
                //p2.Set(flrTh_dbl);


                floorType = floor.FloorType;
                floorType.GetCompoundStructure().SetLayerWidth(0, flrTh_dbl);

                createFloor.Commit();

                return floor;
            }
            #endregion
        }
        #endregion


        #region // DeleteFloor Function to delete floor //

        public void DeleteFloors(Autodesk.Revit.DB.Document actvDoc)
        {
            FilteredElementCollector floorCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(Floor));
            List<Floor> floors = floorCollector.Cast<Floor>().ToList();

            foreach (Floor floor in floors)
            {
                using (Transaction deleteFloors = new Transaction(actvDoc, "Delete Floor"))
                {
                    deleteFloors.Start();
                    actvDoc.Delete(floor.Id);
                    deleteFloors.Commit();
                }
            }
        }

        public void DeleteFloors(Autodesk.Revit.DB.Document actvDoc, List<Floor> floor_lst)
        {
            foreach (Floor floor in floor_lst)
            {
                using (Transaction deleteFloors = new Transaction(actvDoc, "Delete Floor"))
                {
                    deleteFloors.Start();
                    actvDoc.Delete(floor.Id);
                    deleteFloors.Commit();
                }
            }
        }
        #endregion

        #region Delete Dimensions Element ID  List
        public void DeleteDims_lst(Autodesk.Revit.DB.Document actvDoc, List<ElementId> Dims_el_lst)
        {
            foreach (ElementId Dims_el in Dims_el_lst)
            {
                using (Transaction deleteFloors = new Transaction(actvDoc, "Delete Floor"))
                {
                    deleteFloors.Start();
                    actvDoc.Delete(Dims_el);
                    deleteFloors.Commit();
                }
            }
        }
        #endregion



        #region // FindPipeCurve Function to get curve of pipe //
        public Curve FindPipeCurve(Pipe p)
        {
            LocationCurve lc = p.Location as LocationCurve;
            Curve c = lc.Curve;
            XYZ endpoint1 = c.GetEndPoint(0);
            XYZ endpoint2 = c.GetEndPoint(1);
            Curve curve = Line.CreateBound(endpoint1, endpoint2);

            return curve;
        }
        #endregion

        #region // FindFloorFace Function to find faces of floor //
        public List<Face> FindFloorFace(Floor f)
        {
            List<Face> normalFaces = new List<Face>();

            Options opt = new Options
            {
                ComputeReferences = true,
                DetailLevel = ViewDetailLevel.Fine
            };

            GeometryElement e = f.get_Geometry(opt);

            List<Solid> solids = new List<Solid>();

            foreach (GeometryObject obj in e)
            {
                Solid solid = obj as Solid;
                solids.Add(solid);


                if (solid != null && solid.Faces.Size > 0)
                {
                    foreach (Face face in solid.Faces)
                    {
                        PlanarFace pf = face as PlanarFace;
                        if (pf != null)
                        {
                            normalFaces.Add(pf);
                        }
                    }
                }
            }

            return normalFaces;
        }
        #endregion

        #region // FindFaceCurve Function to find intersection of pipe Curve an Floor Face //
        public XYZ FindFaceCurve(Curve pipeCurve, Face floorFace)
        {
            //The intersection point
            IntersectionResultArray intersectionR = new IntersectionResultArray();//Intersection point set

            SetComparisonResult results;//Results of Comparison

            results = floorFace.Intersect(pipeCurve, out intersectionR);

            XYZ intersectionResult = null;//Intersection coordinate

            if (SetComparisonResult.Disjoint != results)
            {
                if (intersectionR != null)
                {
                    if (!intersectionR.IsEmpty)
                    {
                        intersectionResult = intersectionR.get_Item(0).XYZPoint;
                    }
                }
            }
            return intersectionResult;
        }
        #endregion

        #region  // Get Pipe OverallSize Based on:
        //https://thebuildingcoder.typepad.com/blog/2021/10/sci-fi-languages-and-pipe-insulation-retrieval.html#2
        //Fatest get inuslation metthod by 
        //Alexander @aignatovich @CADBIMDeveloper Ignatovich, aka Александр Игнатович,
        double get_PipeOverallSize(Pipe currPipe)
        {
            Autodesk.Revit.DB.Document currPipe_doc = currPipe.Document;
            ElementId currPipe_id = currPipe.Id;
            var pipeInsulation = InsulationLiningBase.GetInsulationIds(currPipe_doc, currPipe_id).Select(currPipe_doc.GetElement).OfType<PipeInsulation>().FirstOrDefault();
            double overallSize_dbl = pipeInsulation?.Thickness ?? 0.0;   // if null set to 0.0
            return (overallSize_dbl * 2.0 + currPipe.Diameter);
        }
        #endregion  // End Of Get Pipe OverallSize Based on

        #region  // Get Pipe insulation Thickness Method 2 --- Fastest--- per:
        //https://thebuildingcoder.typepad.com/blog/2021/10/sci-fi-languages-and-pipe-insulation-retrieval.html#2
        //Alexander @aignatovich @CADBIMDeveloper Ignatovich, aka Александр Игнатович,
        double get_PipeInsulationThickness(Autodesk.Revit.DB.Plumbing.Pipe currPipe)
        {
            // Very Slow Metthod of Gettting Pipe insulation:
            // var pipeInsulation = pipe.GetDependentElements(new ElementClassFilter(typeof(PipeInsulation))).Select(pipe.Document.GetElement).Cast<PipeInsulation>().FirstOrDefault();
            Autodesk.Revit.DB.Document currPipe_doc = currPipe.Document;
            ElementId currPipe_id = currPipe.Id;
            // Much Faster Metthod of Gettting Pipe insulation:
            var pipeInsulation = InsulationLiningBase.GetInsulationIds(currPipe_doc, currPipe_id).Select(currPipe_doc.GetElement).OfType<PipeInsulation>().FirstOrDefault();
            double th = pipeInsulation?.Thickness ?? 0.0;

            return th;
        }
        #endregion  // Get Pipe insulation Thickness Method 2 --- Fastest--- per:

        #region  // Get Pipe insulation Method 2 --- Fastest--- per:
        //https://thebuildingcoder.typepad.com/blog/2021/10/sci-fi-languages-and-pipe-insulation-retrieval.html#2
        //Alexander @aignatovich @CADBIMDeveloper Ignatovich, aka Александр Игнатович,
        PipeInsulation get_PipeInsulation(Autodesk.Revit.DB.Plumbing.Pipe currPipe)
        {
            // Very Slow Metthod of Gettting Pipe insulation:
            // var pipeInsulation = pipe.GetDependentElements(new ElementClassFilter(typeof(PipeInsulation))).Select(pipe.Document.GetElement).Cast<PipeInsulation>().FirstOrDefault();
            Autodesk.Revit.DB.Document currPipe_doc = currPipe.Document;
            ElementId currPipe_id = currPipe.Id;
            // Much Faster Metthod of Gettting Pipe insulation:
            var pipeInsulation = InsulationLiningBase.GetInsulationIds(currPipe_doc, currPipe_id).Select(currPipe_doc.GetElement).OfType<PipeInsulation>().FirstOrDefault();


            return pipeInsulation as PipeInsulation;
        }

        #endregion  // End Of Get Pipe insulation 2 Methods per

        #region  // Get Pipe insulation Method 1 --- Slow ---
        /// <summary>
        /// Return pipe insulation for given pipe using 
        /// filtered element collector and HostElementId
        /// property or InsulationLiningBase 
        /// GetInsulationIds method.
        /// </summary>
        PipeInsulation GetPipeInslationFromPipe(
          Pipe pipe)
        {
            if (pipe == null)
            {
                throw new ArgumentNullException("pipe");
            }

            Autodesk.Revit.DB.Document doc = pipe.Document;

            // Filtered element collector and HostElementId

            FilteredElementCollector fec
              = new FilteredElementCollector(doc)
                .OfClass(typeof(PipeInsulation));

            PipeInsulation pipeInsulation = null;

            foreach (PipeInsulation pi in fec)
            {
                // Find the first insulation
                // belonging to the given pipe

                if (pi.HostElementId == pipe.Id)
                {
                    pipeInsulation = pi;
                    break;
                }
            }

#if DEBUG
            // InsulationLiningBase.GetInsulationIds method
            // returns all pipe insulations for a given pipe

            ICollection<ElementId> pipeInsulationIds
              = InsulationLiningBase.GetInsulationIds(
                doc, pipe.Id);

            Debug.Assert(
              pipeInsulationIds.Contains(pipeInsulation.Id),
              "expected InsulationLiningBase.GetInsulationIds"
              + " to include pipe insulation element id");
#endif // DEBUG

            return pipeInsulation;
        }
        #endregion  // End of Get Pipe insulation Method 1 --- Slow --- 

        #region // IsVertical Function to declare given pipe as vertical or not //
        public bool IsVertical(Pipe p)
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
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
        #endregion


        #region  // Load all Unloaded Documents.... Prefer that user load the documents that has the Floors in it insetead of loading all linked documents.
        public List<RevitLinkType> load_all_Un_LnkdDocs(Autodesk.Revit.DB.Document actvDoc)
        {
            //ISet<ElementId> xrefs = ExternalResourceUtils.GetAllExternalResourceReferences(actvDoc);
            FilteredElementCollector rvtLinks = new FilteredElementCollector(actvDoc).OfCategory(BuiltInCategory.OST_RvtLinks);
            IList<ElementId> rvtLnksElId_lst = rvtLinks.ToElementIds().ToList();
            List<RevitLinkType> loadedDocs_lst = new List<RevitLinkType>();

            string tds = "load_all_Un_LnkdDocs";
            try
            {
                foreach (ElementId eid in rvtLnksElId_lst)
                {
                    var elem = actvDoc.GetElement(eid);
                    if (elem == null) continue;   // Element is not valid... i don't know what this means though... so do nothing and go to next elementid of a rvtLnksElId_lst

                    // Get RVT document links only this time
                    var link = elem as RevitLinkType;
                    if (link == null) continue;  // This means that the element is not a RevitLinkType... possibly a RevitLinnkInstance ... so do nothing and go to next rvtLnksElId_lst
                    try
                    {
                        // Load model temporarily to get the model  path of the cloud link
                        var result = link.Load();
                        loadedDocs_lst.Add(link);
                    }
                    catch (Exception ex) // never catch all exceptions!
                    {
                        TaskDialog.Show(tds, ex.Message);
                    }
                }  // foreach xrefs
            }
            catch (Exception ex)
            {
                TaskDialog.Show(tds, ex.Message);
            }
            return loadedDocs_lst;
        }  // End of load_all_Un_LnkdDocs()
        #endregion   // End Of Load all Unloaded Documents2.... Prefer that user load the documents that has the Floors in it insetead of loading all linked documents.


        #region  // Load all Unloaded Documents.... Prefer that user load the documents that has the Floors in it insetead of loading all linked documents.
        public List<RevitLinkType> load_Selected_Un_LnkdDocs(Autodesk.Revit.DB.Document actvDoc)
        {
            //ISet<ElementId> xrefs = ExternalResourceUtils.GetAllExternalResourceReferences(actvDoc);
            FilteredElementCollector rvtLinks = new FilteredElementCollector(actvDoc).OfCategory(BuiltInCategory.OST_RvtLinks);
            IList<ElementId> rvtLnksElId_lst = rvtLinks.ToElementIds().ToList();
            List<RevitLinkType> loadedDocs_lst = new List<RevitLinkType>();

            string tds = "load_all_Un_LnkdDocs";
            try
            {
                foreach (ElementId eid in rvtLnksElId_lst)
                {
                    var elem = actvDoc.GetElement(eid);
                    if (elem == null) continue;   // Element is not valid... i don't know what this means though... so do nothing and go to next elementid of a rvtLnksElId_lst

                    // Get RVT document links only this time
                    var link = elem as RevitLinkType;
                    if (link == null) continue;  // This means that the element is not a RevitLinkType... possibly a RevitLinnkInstance ... so do nothing and go to next rvtLnksElId_lst
                    DialogResult dialogResult = MessageBox.Show("Do you want to load this Document: " + link.Name, "Loading Linked Documents", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        try
                        {
                            // Load model temporarily to get the model  path of the cloud link
                            var result = link.Load();
                            loadedDocs_lst.Add(link);
                        }
                        catch (Exception ex) // never catch all exceptions!
                        {
                            TaskDialog.Show(tds, ex.Message);
                        }
                    }  //if Load Dialog
                }  // foreach xrefs
            }
            catch (Exception ex)
            {
                TaskDialog.Show(tds, ex.Message);
            }
            return loadedDocs_lst;
        }  // End of load_all_Un_LnkdDocs()
        #endregion   // End Of Load all Unloaded Documents2.... Prefer that user load the documents that has the Floors in it insetead of loading all linked documents.





        #region   // unload all  loaded documents by insertSleeve
        public void unload_loadedDocs_lst(List<RevitLinkType> loadedDocs_lst)
        {
            foreach (RevitLinkType rvtlnktyp in loadedDocs_lst) { rvtlnktyp.Unload(null); }
        }
        #endregion  //End Of unload all  loaded documents by insertSleeve


        #region //  getFlrs_AllLnkdDocs_lst()
        public List<flrsPerLnkdDoc_strct> getFlrs_AllLnkdDocs_lst(Autodesk.Revit.ApplicationServices.Application app, BuiltInCategory BIC_str)
        {
            List<flrsPerLnkdDoc_strct> flrs_kvp_lst = new List<flrsPerLnkdDoc_strct>();
            List<ElementId> flrs_inAllLnkdDocs_lst = new List<ElementId>();
            string tds_lnkd = "---    LINKED DOCS   ---";
            string tds_notlnkd = "---   NOT LINKED DOCS   ---";

            foreach (Autodesk.Revit.DB.Document lnkdDoc in app.Documents)
            {

                if (!lnkdDoc.IsLinked)
                {

                    // how do i load a linked document

                    tds_notlnkd += string.Format("\n - Not Linked document.Title '{0}': ", lnkdDoc.Title);
                    //return retDoc;
                }
                else //(lnkdDoc.IsLinked)
                {

                    tds_lnkd += string.Format("\n - Linked document.Title:  '{0}': ", lnkdDoc.Title);
                    FilteredElementCollector linkedFamCollector = new FilteredElementCollector(lnkdDoc);
                    //ICollection<ElementId> archModelFloors_col = new FilteredElementCollector(lnkdDoc).WhereElementIsNotElementType().OfCategory(BIC_str).ToElementIds();
                    List<ElementId> archModelFloors_lst = new FilteredElementCollector(lnkdDoc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_Floors).ToElementIds().ToList();
                    if (archModelFloors_lst.Count != 0 && archModelFloors_lst != null)
                    {
                        flrsPerLnkdDoc_strct flrsPerLnkdDoc = new flrsPerLnkdDoc_strct();
                        flrsPerLnkdDoc.lnkDoc = lnkdDoc;
                        flrsPerLnkdDoc.flrsElId_lst = archModelFloors_lst;
                        flrs_kvp_lst.Add(flrsPerLnkdDoc);
                    }
                }
            }
            TaskDialog.Show("getFlrs_AllLnkdDocs_lst", tds_lnkd + "\n\n" + tds_notlnkd);








            return flrs_kvp_lst;
        }//  End getFlrs_AllLnkdDocs_lst()
        #endregion  // End getFlrs_AllLnkdDocs_lst()


        #region //  getRvtLnkdDoc()
        public Autodesk.Revit.DB.Document getLnkdDoc(Autodesk.Revit.ApplicationServices.Application app, Autodesk.Revit.DB.Document hostDoc, string matchStr)
        {
            Autodesk.Revit.DB.Document retDoc = null;
            foreach (Autodesk.Revit.DB.Document lnkdDoc in app.Documents)
            {
                if (!lnkdDoc.IsLinked)
                {
                    TaskDialog.Show("getLnkdDoc", string.Format("\n - Not Linked document \n '{0}': ", lnkdDoc.Title));
                    //return retDoc;
                }
                else //(lnkdDoc.IsLinked)
                {
                    TaskDialog.Show("getLnkdDoc", string.Format("\n - Linked document \n '{0}': ", lnkdDoc.Title));
                    retDoc = lnkdDoc;
                }
            }
            return retDoc;
        }//  End getRvtLnkdDoc()
        #endregion  // End getRvtLnkdDoc()


        #region //  getArchModelFloors_Col()
        public ICollection<ElementId> getArchModelFloors_Col(Autodesk.Revit.ApplicationServices.Application app, Autodesk.Revit.DB.Document lnkdDoc, BuiltInCategory BIC_str)
        {
            FilteredElementCollector linkedFamCollector = new FilteredElementCollector(lnkdDoc);
            ICollection<ElementId> archModelFloors_col = new FilteredElementCollector(lnkdDoc).WhereElementIsNotElementType().OfCategory(BIC_str).ToElementIds();

            return archModelFloors_col;
        }//  End getArchModelFloors_Col
        #endregion  // End getArchModelFloors_Col


        #region   //  copyPasteIds
        public static void copyPasteIds(Autodesk.Revit.DB.Document hostDoc, Autodesk.Revit.DB.Document lnkdDoc, IList<ElementId> lnkdFlrs_col)
        {
            //TaskDialog.Show("copyPasteIds", "There Count of Floors in this Doc is: " + lnkdFlrs_col.Count);
            using (Transaction copyPasteLnkdElm_trx = new Transaction(hostDoc, "Copy Elements from Linked Doc and Paste in Host Doc"))
            {
                copyPasteLnkdElm_trx.Start();
                try
                {
                    // // Set failure handler  .. this is to handle me writing the same value in the Mark parameter which i guess, is not expected.
                    FailureHandlingOptions failureOptions = copyPasteLnkdElm_trx.GetFailureHandlingOptions();
                    FailurePreproccessor faliurePreProcessor = new FailurePreproccessor();
                    failureOptions.SetFailuresPreprocessor(faliurePreProcessor);
                    copyPasteLnkdElm_trx.SetFailureHandlingOptions(failureOptions);

                    // Handles copying new types or cancel operation.
                    CopyPasteOptions copyOptions = new CopyPasteOptions();
                    copyOptions.SetDuplicateTypeNamesHandler(new CopyUseDestination());
                    //copyPasteLnkdElm_trx.Start();

                    ElementTransformUtils.CopyElements(lnkdDoc, lnkdFlrs_col, hostDoc, null, copyOptions);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("copyPasteIds", "Exception Copy Paste : " + ex.ToString());
                }
                hostDoc.Regenerate();
                copyPasteLnkdElm_trx.Commit();
            }
        }  // End Of copyPasteIds
        #endregion  // End Of copyPasteIds

        /*  #region   //  copyPasteIdsFromList

          public static void copyPasteIdsFromList(Document hostDoc, List<flrsPerLnkdDoc_strct> flrs_kvp_lst)
          {
              Document lnkdDoc = null;
              List<ElementId> lnkdFlrs_col = new List<ElementId>();

              string tds = "The link contains the specified elements. \n\n\n Count : " + flrs_kvp_lst.Count;

              using (Transaction copyPasteLnkdElm_trx = new Transaction(hostDoc, "Copy Elements from Linked Doc and Past in Host Doc"))
              {
                  foreach (flrsPerLnkdDoc_strct flrs_kvp in flrs_kvp_lst)
                  {
                      lnkdDoc = flrs_kvp.lnkDoc;
                      lnkdFlrs_col = flrs_kvp.flrsElId_lst;
                      tds += "\n - Linked Doc: " + lnkdDoc.Title + ": Floor Count: " + lnkdFlrs_col.Count;
                      foreach (ElementId eid in lnkdFlrs_col) { tds += "\n    - Element ID:" + eid; }
                      copyPasteLnkdElm_trx.Start();
                      try
                      {

                          // // Set failure handler  .. this is to handle me writing the same value in the Mark parameter which i gues, is not expected.
                          var failureOptions = copyPasteLnkdElm_trx.GetFailureHandlingOptions();
                          failureOptions.SetFailuresPreprocessor(new FailurePreproccessor());
                          copyPasteLnkdElm_trx.SetFailureHandlingOptions(failureOptions);

                          // Handles copying new types or cancel operation.
                          CopyPasteOptions copyOptions = new CopyPasteOptions();
                          copyOptions.SetDuplicateTypeNamesHandler(new CopyUseDestination());
                          ElementTransformUtils.CopyElements(lnkdDoc, lnkdFlrs_col, hostDoc, null, copyOptions);
                          hostDoc.Regenerate();
                      }
                      catch (Exception ex)
                      {
                          TaskDialog.Show("copyPasteIdsFromList", "Exception Copy Paste : " + ex.ToString());
                      }

                      copyPasteLnkdElm_trx.Commit();
                  }  // end of for each flrs_kvp_lst
              }  // End of using
              TaskDialog.Show("copyPasteIdsFromList", tds);
          }  // End Of copyPasteIdsFromList
#endregion  // End Of copyPasteIdsFromList
        */
        #region  // CopyUseDestination
        public class CopyUseDestination : IDuplicateTypeNamesHandler
        {
            public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
            {
                return DuplicateTypeAction.UseDestinationTypes;
                //return DuplicateTypeAction.Abort; // UseDestinationTypes;
            }
        }  // End Of CopyUseDestination
        #endregion  // End Of CopyUseDestination

        #region  //insert floors on levels
        public List<Floor> insertLevelFloors(UIApplication app, Autodesk.Revit.DB.Document actvDoc, string flrTh_str)
        {
            List<Floor> levelFloors_lst = new List<Floor>();
            FilteredElementCollector levelCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(Autodesk.Revit.DB.Level));
            List<Autodesk.Revit.DB.Level> levels = levelCollector.Cast<Autodesk.Revit.DB.Level>().ToList();
            //TaskDialog.Show("insertSleeves", "levels.Count: " + levels.Count);

            // Insertion of floor at each level //

            foreach (Autodesk.Revit.DB.Level lvl in levels)
            {
                Floor floor = InsertFloor(app, lvl, flrTh_str);
                //TaskDialog.Show("insertLevelFloors", "after insertFloor" + floor.Name);
                levelFloors_lst.Add(floor);
                //TaskDialog.Show("insertLevelFloors", "after Add(floor)");
            }

            return levelFloors_lst;
        }  // End Of insertLevelFloors
        #endregion  // End of insertLevelFloors

        /*#region  // Function to suppress the "Duplicate Mark Value" Warnings.
        public class ReferencesNotParallelSwallower : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor a)
            {
                var failures = a.GetFailureMessages();
                foreach (var f in failures)
                {
                    var id = f.GetFailureDefinitionId();
                    if (BuiltInFailures.GeneralFailures.DuplicateValue == id)
                    {
                        a.DeleteWarning(f);
                    }
                }
                return FailureProcessingResult.Continue;
            }
        }   // End of DuplicateMarkSwallower
#endregion    // Function to suppress the "Duplicate Mark Value" Warnings.
*/

        #region // Some Swallower.. i would like to implement for pasting floors
        public class FailurePreproccessor : IFailuresPreprocessor
        {
            FailureProcessingResult IFailuresPreprocessor.PreprocessFailures(FailuresAccessor failAccessor)
            {
                int DuplicateValue_cnt = 0;
                int FloorsOverlap_cnt = 0;
                int JoiningDisjoint_cnt = 0;
                int CannotCreateCutout_cnt = 0;
                int DimensionReferencesInvalid_cnt = 0;
                int CannotCutInstanceOutGenError_cnt = 0;
                int ColumnInsideWall_cnt = 0;

                String transactionName = failAccessor.GetTransactionName();
                bool HasErrors = false;
                IList<FailureMessageAccessor> fmsgs = failAccessor.GetFailureMessages();

                if (fmsgs.Count == 0)
                {
                    return FailureProcessingResult.Continue;
                }
                //see if you can make the counter a list of list to hold counter name with it count
                ////if (transactionName.Equals("DeleteLinkedModel"))
                //{
                foreach (FailureMessageAccessor fmsg in fmsgs)
                {
                    FailureSeverity failureSeverity = fmsg.GetSeverity();
                    if (failureSeverity == FailureSeverity.Error) //&& fmsg.GetFailureDefinitionId().Guid.ToString() == "8a9ff20d-fdc2-4f98-87e6-2aa8b71b0c83")
                    {
                        HasErrors = true;
                        TaskDialog.Show("FailureProcessingResult", " Found FailureSeverity.Errors: \n" + fmsg.GetDescriptionText());
                    }

                    var id = fmsg.GetFailureDefinitionId();

                    if (BuiltInFailures.GeneralFailures.DuplicateValue == id)
                    {
                        DuplicateValue_cnt++;
                        failAccessor.DeleteWarning(fmsg);  // DeleteWarning mimics clicking 'Ok' button.
                        //failAccessor.ResolveFailure(fmsg);  // ResolveFailure mimics clicking  'Remove Link' button
                    }   // Elements Have duplicate values.

                    if (BuiltInFailures.OverlapFailures.FloorsOverlap == id)
                    {
                        FloorsOverlap_cnt++;
                        failAccessor.DeleteWarning(fmsg);  // DeleteWarning mimics clicking 'Ok' button.
                        //failAccessor.ResolveFailure(fmsg);  // ResolveFailure mimics clicking  'Remove Link' button
                    }   // Highlighted floors overlap.

                    if (BuiltInFailures.JoinElementsFailures.JoiningDisjoint == id)
                    {
                        JoiningDisjoint_cnt++;
                        //failAccessor.DeleteWarning(fmsg);  // DeleteWarning mimics clicking 'Ok' button.
                        failAccessor.ResolveFailure(fmsg);  // ResolveFailure mimics clicking  'Remove Link' button
                    }  // Highlighted Elements are joined but do Not intersect.

                    if (BuiltInFailures.CreationFailures.CannotCreateCutout == id)
                    {
                        CannotCreateCutout_cnt++;
                        failAccessor.DeleteWarning(fmsg);  // DeleteWarning mimics clicking 'Ok' button.
                        //failAccessor.ResolveFailure(fmsg);  // ResolveFailure mimics clicking  'Delete Elements' button
                    }   // Can't make cut-out

                    if (BuiltInFailures.DimensionFailures.DimensionReferencesInvalid == id)
                    {
                        DimensionReferencesInvalid_cnt++;
                        failAccessor.DeleteWarning(fmsg);  // DeleteWarning mimics clicking 'Ok' button.
                        //failAccessor.ResolveFailure(fmsg);  // ResolveFailure mimics clicking  'Delete Dimension References' button
                    }   // One or more dimension references are or have become invalid.

                    if (BuiltInFailures.CutFailures.CannotCutInstanceOutGenError == id)
                    {
                        CannotCutInstanceOutGenError_cnt++;
                        //failAccessor.DeleteWarning(fmsg);  // DeleteWarning mimics clicking 'Ok' button.
                        failAccessor.ResolveFailure(fmsg);  // ResolveFailure mimics clicking  'Delete Instances' button
                    }   // Can't cut instances of [symbol] out of its host.

                    if (BuiltInFailures.ColumnInsideWallFailures.ColumnInsideWall == id)
                    {
                        ColumnInsideWall_cnt++;
                        //failAccessor.DeleteWarning(fmsg);  // DeleteWarning mimics clicking 'Ok' button.
                        failAccessor.ResolveFailure(fmsg);  // ResolveFailure mimics clicking  'Delete Instances' button
                    }   // One Element is completely inside another
                }

                //}  // if Transaction name is...
                /*TaskDialog.Show("FailureProcessingResult", " \n Failures \n " +
                    "\n Dupicates:       " + dupes_cnt +
                    "\n FloorsOverlap:   " + flrsOvrlp_cnt +
                    "\n JoiningDisjoint: " + jnDsjnt_cnt
                    );*/

                if (!HasErrors)
                {
                    return FailureProcessingResult.Continue;
                }

                //failAccessor.DeleteAllWarnings();     // Delete ALL warnings
                //failAccessor.ResolveFailures(fmsgs);  //  Resolves ALL Failures
                return FailureProcessingResult.ProceedWithCommit;
            }
        }
        #endregion //End of  Some Swallower.. i would like to implement for pasting floors


        #region // Find if XYZ is in a List of XYZ with a tolerance using IEnumerale 
        // Check if an xyz point is "found" in a some list of points. "found" here means within +/- xyz_rng.
        // The xyz_rng cannot have 0 in it otherwise autodesk will return false results, since it is comparing decimal numbers, e.g. 0.33333333

        public bool isWithinRange(List<XYZ> xyz_lst, XYZ xyz_pt, XYZ xyz_rng)
        {
            //TaskDialog.Show("isWithinRange", "in The isWithinRange: ");
            IEnumerable<XYZ> xyz_ien = xyz_lst as IEnumerable<XYZ>;

            List<XYZ> SleevesIntersects_lst = xyz_ien.Where(loc =>
            (loc.X > xyz_pt.X - xyz_rng.X && loc.X < xyz_pt.X + xyz_rng.X) &&
            (loc.Y > xyz_pt.Y - xyz_rng.Y && loc.Y < xyz_pt.Y + xyz_rng.Y) &&
            (loc.Z > xyz_pt.Z - xyz_rng.Z && loc.Z < xyz_pt.Z + xyz_rng.Z)).ToList<XYZ>();

            if (SleevesIntersects_lst.Count > 0) { return true; }

            else return false;
        }




        #endregion



        #region  // OPens fileDialog Box for File selection
        public void openFileDialogBox(Autodesk.Revit.DB.Document actvDoc)
        {
            // Get application and document objects

            //UIDocument uidoc = ActiveUIDocument;
            //Autodesk.Revit.ApplicationServices.Application app = ThisApplication.Application;
            //Autodesk.Revit.DB.Document doc = uidoc.Document;

            try
            {
                using (Transaction transaction = new Transaction(actvDoc))
                {
                    // Link files in folder
                    transaction.Start("Link files");

                    OpenFileDialog openFileDialog1 = new OpenFileDialog();
                    openFileDialog1.InitialDirectory = (@"P:\");
                    openFileDialog1.Filter = "RVT|*.rvt";
                    openFileDialog1.Multiselect = true;
                    openFileDialog1.RestoreDirectory = true;

                    if (openFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        // string[] filesInFolder = openFileDialog1.FileNames;
                        foreach (string path in openFileDialog1.FileNames)
                        {
                            FileInfo filePath = new FileInfo(path);

                            // debug ***********
                            TaskDialog.Show("insertSleeves", "filePath.FullName.ToString() = " + filePath.FullName.ToString());
                            // debug ***********

                            ModelPath linkpath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath.FullName.ToString());
                            RevitLinkOptions options = new RevitLinkOptions(false);
                            LinkLoadResult result = RevitLinkType.Create(actvDoc, linkpath, options);
                            RevitLinkInstance.Create(actvDoc, result.ElementId);
                        }
                    }
                    // Show summary message
                    //TaskDialog.Show("Files", filePaths.ToString());
                    // Assuming that everything went right return Result.Succeeded
                    transaction.Commit();
                    //return Result.Succeeded;
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // If user decided to cancel the operation return Result.Canceled
                TaskDialog.Show("test", "User Canceled ");
            }
            catch (Exception ex)
            {
                // If something went wrong return Result.Failed
                Console.WriteLine("There was a problem!");
                Console.WriteLine(ex.Message);
                //return Result.Failed;
            }
        }  // End openFileDialog

        #endregion  // End OPens fileDialog Box for File selection

        #region // flrsPerLnkdDoc_strct
        public struct flrsPerLnkdDoc_strct
        {
            public Autodesk.Revit.DB.Document lnkDoc;
            public List<ElementId> flrsElId_lst;

            public flrsPerLnkdDoc_strct(Autodesk.Revit.DB.Document lnkDoc, List<ElementId> flrsElId_lst)
            {
                this.lnkDoc = lnkDoc;
                this.flrsElId_lst = flrsElId_lst;

            }

        } // End of flrsPerLnkdDoc_strct
        #endregion  // End of flrsPerLnkdDoc_strct



    }  // End Of Class InsertSleeve 

    #region Wr dialog class
    public static class dialogInput
    {
        public static T ReadLine<T>(string message)
        {
            Console.WriteLine(message);
            string input = Console.ReadLine();
            return (T)Convert.ChangeType(input, typeof(T));
        }
    }
    #endregion  End Of Dialog Class





}   // End of Namespace KDS_Module

