using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;




namespace KDS_Module_Vx
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class InsertSleeve : IExternalCommand
    {
        const string _sleeveString = "KDS_Hilti-FS_CFS_CID";
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            #region // Some Definitions 
            Document actvDoc = commandData.Application.ActiveUIDocument.Document;
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            UIApplication uiApp = commandData.Application;
            Autodesk.Revit.ApplicationServices.Application app = uiApp.Application;
            List<Point> sleevesLocPnt_lst = new List<Point>();
            #endregion  // Some Definitions

            #region   // Get Family of Sleeves and List of all Pipes in Host Doc.sleevesLocPnt_lst
            // Get The FamilySymbol for the Sleeve
            FilteredElementCollector sleeveCollector = new FilteredElementCollector(actvDoc);
            Family sleeveFam = sleeveCollector.OfClass(typeof(Family)).OfType<Family>().FirstOrDefault(f => f.Name.Equals(_sleeveString));

            FamilySymbol sleeveFamSymbol = actvDoc.GetElement(sleeveFam.GetFamilySymbolIds().FirstOrDefault()) as FamilySymbol;


            // Get a List of All Pipes that are Vertical and within 2" and 6" size
            FilteredElementCollector pipeCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(Pipe));
            List<Pipe> pipes_lst = pipeCollector.Cast<Pipe>().ToList();
            List<Pipe> pipes_size_lst = pipes_lst.Where(p => p.Diameter > 0.146 && p.Diameter < .667).ToList<Pipe>();
            List<Pipe> pipes_size_slope_lst = pipes_size_lst.Where(p => IsVertical(p)).ToList<Pipe>();
            TaskDialog.Show("insertSleeve", "All Available Pipes Count: " + pipes_lst.Count +
               "\n Pipes of size between 2 and 6 inches Count: " + pipes_size_lst.Count +
               "\n Pipes of correct size and Sloped vertically Count: " + pipes_size_slope_lst.Count);
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
                insertSleeves(commandData.Application, pipes_size_slope_lst, hostDoc_floors_lst, sleeveFamSymbol);

            }
            else   // No Floors in Host Doc, so Check the Linked Models
            {


                DialogResult dialogResult = MessageBox.Show("Check Linked Documents?", "Insert Sleeves", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    List<RevitLinkType> loadedDocsLinkTypes_lst = new List<RevitLinkType>();
                    loadedDocsLinkTypes_lst = load_all_Un_LnkdDocs(actvDoc);

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
                        int fc = 0;
                        tds += " \n " + dc + "- Floors in : " + flrs_kvp.lnkDoc.Title;
                        foreach (ElementId feid in flrs_kvp.flrsElId_lst)
                        {
                            fc++;
                            tf++;
                            Floor f = flrs_kvp.lnkDoc.GetElement(feid) as Floor;
                            tds += "\n   " + fc + "- Floor Name: " + f.Name;
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

                    insertSleeves(commandData.Application, pipes_size_slope_lst, hostDoc_cpdFlrs_lst, sleeveFamSymbol);


                    TaskDialog.Show("insertSleeves", " Deleting All Copyed Floors from Linked Documents.");
                    DeleteFloors(actvDoc, hostDoc_cpdFlrs_lst);
                }

                else   // Neither Host Doc, Nor Linked Docs have Floors, so create ones based on host Doc Levels and get instersections.
                {
                    TaskDialog.Show("insertSleeves", "    Neither Host Doc, Nor Linked Docs have Floors defined.  \n - Create Temp Floors based on host Doc Levels. \n - Place Sleeves at instersections with pipes. \n - Delete Temp Floors");
                    List<Floor> levelsFloors_lst = new List<Floor>();
                    // Create Temp Floors in Host DOc
                    levelsFloors_lst = insertLevelFloors(commandData.Application, actvDoc);
                    TaskDialog.Show("insertSleeves", "Created new Temp Floors. \n levelsFloors_lst.Count = " + levelsFloors_lst.Count);
                    // Insert Floor Sleeves ata intersections
                    insertSleeves(commandData.Application, pipes_size_slope_lst, levelsFloors_lst, sleeveFamSymbol);

                    TaskDialog.Show("insertSleeves", " Delete Temp Floors");
                    // Delete Temp Floors in Host DOc
                    DeleteFloors(actvDoc);
                }


            }

            #endregion

            return Result.Succeeded;
        }


        #region // insertSleeves Function to place sleeve where Pipe and Level intersect --- No Floor exists here. so we need to create floors as well//
        public List<XYZ> insertSleeves(UIApplication uiApp, List<Pipe> pipes_lst, List<Floor> floors_lst, FamilySymbol sleeveFamSymbol)
        {
            #region // Defining and Collection of Elements and Variables //

            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document actvDoc = uidoc.Document;


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
                                    // // Set failure handler  .. this is to handle me writing the same value in the Mark parameter which i gues, is not expected.
                                    var failureOptions = setSleeveParam_trx.GetFailureHandlingOptions();
                                    failureOptions.SetFailuresPreprocessor(new FailurePreproccessor());
                                    setSleeveParam_trx.SetFailureHandlingOptions(failureOptions);

                                    //TaskDialog.Show("insertSleeve", " Just Disabled Warnings on Duplicate Mark Values");

                                    //Parameter prm = _p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                                    double p_diam = p.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM).AsDouble();
                                    if (p_diam >= 0.41665 && p_diam <= .41668) { p_diam = 0.5; }
                                    //TaskDialog.Show("insertSleeve", " p_diam: " + p_diam.ToString());

                                    // Get the Diameter Parameter for the Sleeve THEN set it to the size from the relevant pipe.
                                    Parameter slv_diam = sleeve_famInst.LookupParameter("Size");
                                    slv_diam.Set(p_diam + (1.0 / 6.0));

                                    //TaskDialog.Show("insertSleeve", " slv_diam: " + slv_diam.AsValueString());

                                    // Get the System Type  Parameter for the Sleeve THEN set it to the size from the relevant pipe.
                                    Parameter p_SysType = p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                                    //TaskDialog.Show("insertSleeve", " p_SysType: " + p_SysType.AsValueString());

                                    //Parameter slv_SysType = sleeve_famInst.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM);
                                    Parameter slv_Mark = sleeve_famInst.LookupParameter("Mark");

                                    //TaskDialog.Show("insrtSleeve", "\n BEFORE set Mark: " + slv_Mark.AsValueString()   );

                                    //slv_SysType.Set(p_SysType.AsValueString());
                                    slv_Mark.Set(p_SysType.AsValueString());

                                    //TaskDialog.Show("insrtSleeve","AFTER set Mark: " + sleeve_famInst.LookupParameter("Mark").AsValueString());
                                }
                                catch (Exception e)
                                {

                                    TaskDialog.Show("insrtSleeve", "  In insertSleeves Function " + "\nException try to write parameters: e: " + e);
                                }
                                setSleeveParam_trx.Commit();
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


        #region // Reset the shape of the floors as Elementid 
        public void resetFloor(Document actvDoc, List<ElementId> floors_lst)
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

        public void GetInfo_FloorType(FloorType floorType)
        {
            string message;
            // Get whether FloorType is a foundation slab
            message = "If is foundation slab : " + floorType.IsFoundationSlab;
            TaskDialog.Show("Revit", message);
        }


        #region // Reset the shape of the floors as Floor 
        public void resetFloor(Document actvDoc, List<Floor> floors_lst)
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
        public Floor InsertFloor(UIApplication uiApp, Level lvl)
        {
            UIDocument uidoc = uiApp.ActiveUIDocument;
            Document actvDoc = uidoc.Document;

            Floor floor = null;
            #region // Transaction to Create Floor //
            using (Transaction createFloor = new Transaction(actvDoc, "Create floor"))
            {
                createFloor.Start();
#if RVT2020
                #region // Get a floor type for floor creation 2020
                FloorType floorType = new FilteredElementCollector(actvDoc).OfClass(typeof(FloorType)).First(e => e.Name.Equals("Generic - 12\"")) as FloorType;   //For Revit 2021 and older
      
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
                FloorType floorType = new FilteredElementCollector(actvDoc).OfClass(typeof(FloorType)).First(e => e.Name.Equals("Generic - 12\"")) as FloorType;   //For Revit 2021 and older
      
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

                p.Set(lvl.Id);
                p1.Set(0);

                createFloor.Commit();
                return floor;
            }
#endregion
        }
#endregion


#region // DeleteFloor Function to delete floor //

        public void DeleteFloors(Document actvDoc)
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

        public void DeleteFloors(Document actvDoc, List<Floor> floor_lst)
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
        public void DeleteDims_lst(Document actvDoc, List<ElementId> Dims_el_lst)
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
        public List<RevitLinkType> load_all_Un_LnkdDocs(Document actvDoc)
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

            foreach (Document lnkdDoc in app.Documents)
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
        public Document getLnkdDoc(Autodesk.Revit.ApplicationServices.Application app, Document hostDoc, string matchStr)
        {
            Document retDoc = null;
            foreach (Document lnkdDoc in app.Documents)
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
        public ICollection<ElementId> getArchModelFloors_Col(Autodesk.Revit.ApplicationServices.Application app, Document lnkdDoc, BuiltInCategory BIC_str)
        {
            FilteredElementCollector linkedFamCollector = new FilteredElementCollector(lnkdDoc);
            ICollection<ElementId> archModelFloors_col = new FilteredElementCollector(lnkdDoc).WhereElementIsNotElementType().OfCategory(BIC_str).ToElementIds();

            return archModelFloors_col;
        }//  End getArchModelFloors_Col
#endregion  // End getArchModelFloors_Col


#region   //  copyPasteIds
        public static void copyPasteIds(Document hostDoc, Document lnkdDoc, IList<ElementId> lnkdFlrs_col)
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
        public List<Floor> insertLevelFloors(UIApplication app, Document actvDoc)
        {
            List<Floor> levelFloors_lst = new List<Floor>();
            FilteredElementCollector levelCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(Level));
            List<Level> levels = levelCollector.Cast<Level>().ToList();
            //TaskDialog.Show("insertSleeves", "levels.Count: " + levels.Count);

            // Insertion of floor at each level //

            foreach (Level lvl in levels)
            {
                Floor floor = InsertFloor(app, lvl);
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
        // Check if and xyz point is in a list of points.
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
        public void openFileDialogBox(Document actvDoc)
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
            public Document lnkDoc;
            public List<ElementId> flrsElId_lst;

            public flrsPerLnkdDoc_strct(Document lnkDoc, List<ElementId> flrsElId_lst)
            {
                this.lnkDoc = lnkDoc;
                this.flrsElId_lst = flrsElId_lst;

            }

        } // End of flrsPerLnkdDoc_strct
#endregion  // End of flrsPerLnkdDoc_strct



    }  // End Of Class InsertSleeve 
}   // End of Namespace KDS_Module

