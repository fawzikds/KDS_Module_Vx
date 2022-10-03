/*
 * Created by SharpDevelop.
 * User: cad16
 * Date: 1/4/2021
 * Time: 11:26 AM
 */
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System.Collections.Generic;
using System.Linq;





namespace KDS_Module
{

    /// <summary>
    /// Description of rplcFamilyInstances_CN finds all familyInstances of certain Name, and  certain value of their  "Angle 1" parameter, then replace with an equivalent familyInstance.
    /// 45 degrees  use a wye.
    /// 90degrees use a Combo..
    /// </summary>

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class rplcFamilyInstances_CN : IExternalEventHandler
    {
        #region  //No Code// Replacing Tee with a wye is not easy in Revit since newTeeFitting function does not work for 45 degrees angles.
        /*
        There are 2 work arounds that i know of. 
        I- A work around for using the newTeefitting:
        1. Create branch pipe with the right angle(if I try to create with 45 degrees, NewTeeFitting throws exception "too smal or too large angle")
        2. call NewTeeFitting to create the fitting.
        3. Disconnect the branch pipe from the fitting (if rotate the pipe without disconnecting, commit transaction give the message and disconnect itself)
        4. Rotate the branch pipe
        5. Connect back the branch pipe to the fitting(connection is done, but is not visual effect in drawing)
        6. To get the visual effect in drawing, change fitting type to any type different from required and change again back to required type.

        II- Use Pipe Placeholders
        1- Based on the connectors of a Tee fitting, get their opposing pipe connectors
        2- Create the pipes of the opposing connectors
        3- Create pipe placeholders out of these pipes
        4- find which placeholder pipes are linear (180 degrees between them) then join them into one
        5- create a wye by simply joining the linear pipeholders with the branch pipeholder. ConnectPipePlaceholdersAtTee
        6- Transform the placeholders into a actual pipe. ConvertPipePlaceholders


        Currently placeholder method is still not functional.



        1- Get the following from original Fitting
         - Connector List 
         - Center  
         - Transform
         - Level
        2- Delete original Fitting
        3- Insert a replacement Fitting at same Center
        4- Rotate replacement  fitting in original fitting orientation
        5- Set Replacement fitting size to a default value.
        6- Connect and rezise the fittings to pipe or falnges, whathever the opposing fittings happen to be off.

        */
        #endregion  // Replacing Tee with a wye is not easy in Revit since newTeeFitting function does not work for 45 degrees angles.

        public void Execute(UIApplication a)
        {

            UIDocument uiDoc = a.ActiveUIDocument;

            Document actvDoc = uiDoc.Document;


            #region  // Define Family Names for collector search.
            /*
            string selectString = "Tee - Welded - Generic";
            const string param_name = "Angle";
            string desiredAngleValue = "135";
            string replaceString = "Wye - Welded - Generic";
            */

            // It looks like for KDS_Char_CI_NH, i need to use the Wye as default for designer, and then later replace with Combos.
            // Since using San Tee as Default adds a 45 degrees to make a wierd fitting combination for a Wye.

            string selectString = "KDS_Char_CI_NH_Wye";
            const string param_name = "Angle";
            string desiredAngleValue = "90";
            string replaceString = "KDS_Char_CI_NH_Combo";
            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "Finished - Define Family Names for collector search!");

            #endregion  // Define Family Names for collector search.

            #region // Get the replacement Family based on its name.
            FilteredElementCollector collector = null;
            collector = new FilteredElementCollector(actvDoc);
            Family rplcFam = null;
            rplcFam = collector.OfClass(typeof(Family)).OfType<Family>().FirstOrDefault(f => f.Name.Equals(replaceString));

            FamilySymbol rplcFamSymbol = null;
            rplcFamSymbol = actvDoc.GetElement(rplcFam.GetFamilySymbolIds().FirstOrDefault()) as FamilySymbol;
            rplcFamSymbol.Activate();   // This needs to be in a transaction or else it will result with an exception error.
                                        //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", string.Format("Finished - Get the replacement Family  based on its name: \n {0} \n with Symbol Name: \n {1}", rplcFam.Name, rplcFamSymbol.Name));


            #endregion //Get the replacement Family  based on its name. 

            #region // Get All Family instances to be replaced in Project (FamilyName = "Tee - Weld - Generic"  and param_name = to desiredAngleValue  "Angle 1" = 45)
            collector = new FilteredElementCollector(actvDoc);
            List<FamilyInstance> allFamInst = collector
                             .OfClass(typeof(FamilyInstance))
                             .Cast<FamilyInstance>()
                             .Where(inst => inst.Symbol.FamilyName.Equals(selectString) && inst.LookupParameter(param_name).AsValueString().Contains(desiredAngleValue))
                             .ToList();

            // for Debug purposes.
            var allFamInstCount = allFamInst.Count;
            Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "Finished - Get All Family instances to be replaced \n allFamInstCount = :  " + allFamInstCount);
            #endregion // Get All Family instances to be replaced in Project (FamilyName = "Tee - Weld - Generic" and param_name = to desiredAngleValue  "Angle 1" = 45)


            #region // Main Loop Itirate over all Found instance of Found Families and replace with desired families 
            foreach (Autodesk.Revit.DB.FamilyInstance famInst in allFamInst)
            {

                newTeeFitting_kds(famInst, rplcFamSymbol);
            }  // Foreach Family to edit

            Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in rplcFamilies_CN.execute: \n \n COMPLETED !! \n\n");
            //actvDoc.Regenerate();    // This undoes everything that the previous transactions has comitted !! why?  I was hoping it will re-attach/align  fittings
            //uiDoc.RefreshActiveView();  // This does not change anything.

        }// end of Execute


        public string GetName()
        {
            return "External Event rplcFamilyInstances_CN";
        }


        public FamilyInstance newTeeFitting_kds(FamilyInstance famInst, FamilySymbol rplcFamSymbol)
        {

            Document actvDoc = famInst.Document;

            FamilyInstance rplc_famInst;

            try
            {
                #region  // Get Transform of famInst and Store in a Matrix... It will be used later to re-orient the replacement family instance
                Transform famInst_transform = famInst.GetTransform();

                double[][] famInst_matrix = new double[][]
                {
                        new double[]{ famInst_transform.BasisX.X, famInst_transform.BasisY.X, famInst_transform.BasisZ.X },
                        new double[]{ famInst_transform.BasisX.Y, famInst_transform.BasisY.Y, famInst_transform.BasisZ.Y },
                        new double[]{ famInst_transform.BasisX.Z, famInst_transform.BasisY.Z, famInst_transform.BasisZ.Z }
                };

                #endregion  // Get Transform of famInst

                #region // Get Level and Center from original famInst to use later in Insert replacement Family
                // Get Center of Fitting to be used later.
                XYZ rplcFamSymbol_center = ((LocationPoint)famInst?.Location)?.Point;
                // Get the Level of famInst
                Level famInst_refLevel = actvDoc.GetElement(famInst.LevelId) as Level;
                #endregion // Get Level from original famInst to use later in Insert replacement Family

                #region  // Get Connectors and Opposing connecotrs of famInst...
                List<Connector> famInst_conn_lst = new List<Connector>();
                List<Connector> famInst_opConn_lst = new List<Connector>();  // Connector list to hold FamInst Connectors, regardelss of what they are for, Flange or Pipe.

                // Get all connectors of FamInst
                famInst_conn_lst = famInst.MEPModel.ConnectorManager.Connectors.Cast<Connector>()
                                  .Where(conn => conn.ConnectorType == ConnectorType.End || conn.ConnectorType == ConnectorType.Curve || conn.ConnectorType == ConnectorType.Physical)
                                  .ToList();

                //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "\n\n Finished Getting famInst Connectors\n\n famInst_conn_lst.Count: " + famInst_conn_lst.Count);

                // Get all opposing connectors for the famInst Fitting... These are the connectors on the Pipes or the Flanges connected to famInst, and will be connected to rplc_famInst.. Exclude NUlls
                for (int i = 0; i < famInst_conn_lst.Count; i++)
                {
                    Connector tmpConn;
                    if (famInst_conn_lst[i] != null)
                    {
                        //GetElementAtConnector(famInst_conn_lst[i]);  // Informational
                        // look for physical connections
                        if (famInst_conn_lst[i].ConnectorType == ConnectorType.End ||
                            famInst_conn_lst[i].ConnectorType == ConnectorType.Curve ||
                            famInst_conn_lst[i].ConnectorType == ConnectorType.Physical)
                        {
                            tmpConn = getOpposingPipeConnector(famInst_conn_lst[i]);
                            if (tmpConn != null) { famInst_opConn_lst.Add(tmpConn); }
                        }
                    }
                }
                //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "\n\n Finished Getting Opposing Connectors\n\n famInst_opConn_lst.Count: " + famInst_opConn_lst.Count);
                #endregion //  Get Connectors and Opposing connecotrs of famInst...

                #region  // Delete Original Family Instance: famInst
                using (Transaction delFamInst_trx = new Transaction(actvDoc, "Delete Original Family Instance: famInst"))
                {
                    delFamInst_trx.Start();
                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "Deleting Original Family Instance: famInst \n facingOrientation: " + famInst.FacingOrientation.ToString() +
                    //    "\n " + famInst.HandOrientation.ToString());
                    actvDoc.Delete(famInst.Id);
                    delFamInst_trx.Commit();
                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "Deleted Original Family Instance: famInst");
                }
                #endregion // Delete Original Family Instance: famInst

                #region // Insert a Replace Family Instance
                using (Transaction rplc_famInst_trx = new Transaction(actvDoc, "Replace Family Instance"))
                {
                    rplc_famInst_trx.Start();

                    // Using newTeeFitting, results in an exception unable to insert tee.
                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "\n\n Using newTeeFitting !!\n\n ");
                    //rplc_famInst = actvDoc.Create.NewTeeFitting(famInst_opConn_lst[0], famInst_opConn_lst[1], famInst_opConn_lst[2]);//This does not work for flange system, since connector owners should be a pipe or duct, not a flange. Moreover, if any of these is null, it is another can of worms.

                    rplc_famInst = actvDoc.Create.NewFamilyInstance(rplcFamSymbol_center, rplcFamSymbol, famInst_refLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                    string facingOrientation_txt = rplc_famInst.FacingOrientation.ToString();
                    rplc_famInst_trx.Commit();
                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "Finished Trx: rplc_famInst_trx : Replacement Family Instance: rplc_famInst\n facingOrientation: " + rplc_famInst.FacingOrientation.ToString() +
                    //    "\n " + rplc_famInst.HandOrientation.ToString() + "\n Name" +
                    //    "\n\n rplc_famInst.Symbol.FamilyName: " + rplc_famInst.Symbol.FamilyName +
                    //    "\n rplc_famInst.Name: " + rplc_famInst.Name);
                }//end transaction rplc_famInst_trx
                #endregion // Insert a Replace Family Instance

                #region // Rotate rplcFam_inst based on Matrix

                //  for a wye, i can 
                // Find colinear pipes, pipes that would form a line 
                // in the fitting find the connectors that form a line.
                // find the angle between the two lines
                // rotate the fitting by that angle so they line up.
                //checked third connector wi the other 3rd connector. if same angle and opposite do nothing
                //        if same angle but same direction thenrotate 180 degrees
                //    if 90 degrees flip around an axxis perpendicular to the colinear lines
                //    if 270 flip around the colinear axis.

                //    but what about a cross.?


                using (Transaction rotate_rplc_famInst_trx = new Transaction(actvDoc, "Rotate rplcFam_inst based on Matrix"))
                {
                    rotate_rplc_famInst_trx.Start();
                    var anAx = GetAxisAngleFromMatrix(famInst_matrix);
                    double angleOfRotation = anAx.Item1;
                    XYZ axisOfRotation = anAx.Item2;

                    Line rotationLine = Line.CreateUnbound(famInst_transform.Origin, axisOfRotation);
                    ElementTransformUtils.RotateElement(actvDoc, rplc_famInst.Id, rotationLine, angleOfRotation);
                    rotate_rplc_famInst_trx.Commit();
                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "Finished Trx: rotate_rplc_famInst_trx : Rotate rplcFam_inst based on Matrix. \n anAx = " + anAx);
                }
                #endregion // Rotate rplcFam_inst based on Matrix

                #region // Set non-connected ports in the replacement fitting to some known size
                // I am doing this to circumvent the case where a port is not connected to anything, not even a flange...see below.
                foreach (Connector r_conn in rplc_famInst.MEPModel.ConnectorManager.Connectors)
                {
                    using (Transaction setRadius_trx = new Transaction(actvDoc, "Set Radius to Non-Connected POrts of Replacement Famitting"))
                    {
                        setRadius_trx.Start();
                        r_conn.Radius = 0.125;
                        setRadius_trx.Commit();
                    }
                }
                #endregion // Set non-connected ports in the replacement fitting to some known size

                #region // Connect the Replacement Family Instance to pipes Adding Flanges if required
                Connector rplc_Conn_tmp;

                // Loop Through connectors of Replacement Family (rplc_famInst) and find which is a pipe or a flange/fitting
                //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "\n\n Looping thru famInst Opposing Connectors list.\n\n famInst_opConn_lst.Count = " + famInst_opConn_lst.Count);
                for (int i = 0; i < famInst_opConn_lst.Count; i++)
                {
                    // Check Type of currrent connector's Owner
                    Element pipeFlange_el = actvDoc.GetElement(famInst_opConn_lst[i].Owner.Id);
                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "\n\n I am in Looping thru famInst Opposing Connectors list.. \n\n i = " + i + "\n\n pipeFlange_el id = " + pipeFlange_el.Id);
                    if (pipeFlange_el is Autodesk.Revit.DB.Plumbing.Pipe)    // If Element is Pipe
                    {
                        using (Transaction connPipe_trx = new Transaction(actvDoc, "Connect the Replacement Family Instance to Pipe; \n Set appropriate diameter to fitting port"))
                        {
                            connPipe_trx.Start();
                            rplc_Conn_tmp = connectTwoConn_basedOnAngleOnPrevFitting(famInst_opConn_lst[i], rplc_famInst);
                            connPipe_trx.Commit();      //  This connects the fitting as a whole even when just one port is connected.  but the size is off.. so technically, i do not need to continue with the loop.
                                                        //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "Finished Trx: connPipe_trx: Connected connector at: \n\n i= " + i + "\n\n Pipe id = " + pipeFlange_el.Id);
                        }
                        if (rplc_Conn_tmp != null)
                        {
                            resizeFittingConnPipeFlange(famInst_opConn_lst[i], rplc_Conn_tmp);
                            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "Finished Trx: reszFitt2Pipe_trx: ");
                        }
                        else
                        {
                            Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "Could not Connect Pipe connector at: \n i= " + i);
                        }
                    }
                    else
                    {
                        using (Transaction connFlange_trx = new Transaction(actvDoc, "Connect the Replacement Family Instance to Flanges; \n Set appropriate diameter to fitting port"))
                        {
                            connFlange_trx.Start();
                            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "Started Trx: connFlange_trx: Connected connector at: \n\n i= " + i + "\n\n Flange id = " + pipeFlange_el.Id);
                            rplc_Conn_tmp = connectTwoConn_basedOnAngleOnPrevFitting(famInst_opConn_lst[i], rplc_famInst);
                            connFlange_trx.Commit();
                            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "Finished Trx: connFlange_trx: Connected connector at: \n\n i= " + i + "\n\n Flange id = " + pipeFlange_el.Id);
                        }
                        if (rplc_Conn_tmp != null)
                        {
                            resizeFittingConnPipeFlange(famInst_opConn_lst[i], rplc_Conn_tmp);
                            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "Finished Trx: connFlange_trx: Resize fitting to Flange connector");
                        }
                        else
                        {
                            Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "Could not Connect Flange connector at: \n i= " + i);
                        }
                    }
                }
                #endregion // Connect the Replacement Family Instance to pipes Adding Flanges if exist

                #region  // Nudge Fitting to make force revit to reacttch the pipe properly... currently there is missalignments and gaps.
                using (Transaction nudgeFitting_trx = new Transaction(actvDoc, "Nudge Fitting to make force revit to reacttch the pipe properly"))
                {
                    nudgeFitting_trx.Start();
                    nudgeFitting(rplc_famInst);
                    nudgeFitting_trx.Commit();      //  This connects the fitting as a whole even when just one port is connected.  but the size is off.. so technically, i do not need to continue with the loop.
                                                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "Finished Trx: nudgeFitting_trx: nudged fitting : rplc_famInst id = " + rplc_famInst.Id);
                }

                #endregion    // Nudge Fitting to make force revit to reacttch the pipe properly... currently there is missalignments and gaps.


                //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in rplcFamilies_CN.execute: \n \n Connect the Replacement Family Instance to Pipe; \n Set appropriate diameter to fitting port");

                #endregion // Main Loop Itirate over all Found instance of Found Families and replace with desired families
                return rplc_famInst;
            }//end try
            catch (System.Exception e)
            {
                Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in rplcFamilies_CN.execute. Exception:  " + e.ToString());

            }// end catch
             //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in rplcFamilies_CN.execute: \n \n COMPLETED Main Loop!! \n\n");
            return null;
        }


        #region  // Function: Nudge Fitting to make force revit to reacttch the pipe properly... currently there is missalignments and gaps.
        public void nudgeFitting(FamilyInstance rplc_famInst)
        {
            //XYZ rplcFamSymbol_center = ((LocationPoint)rplc_famInst?.Location)?.Point;   // I need to calculate the center anyways, since it may have moved during connection to pipe.

            LocationPoint rplcFamSymbol_Loc = rplc_famInst.Location as LocationPoint;

            XYZ rplcFamSymbol_cntr = rplcFamSymbol_Loc.Point;

            //  Move the Fitting to new location... by so persentation of previous location... here 0.01%
            XYZ newPlace_diff = new XYZ(0.0001 * rplcFamSymbol_cntr.X, 0.0001 * rplcFamSymbol_cntr.Y, 0.0001 * rplcFamSymbol_cntr.Z);
            rplc_famInst.Location.Move(newPlace_diff);
            // Move bacvk fitting to original place.,.. so negative the previous move.
            XYZ newPlace_diff_bk = new XYZ(-0.0001 * rplcFamSymbol_cntr.X, -0.0001 * rplcFamSymbol_cntr.Y, -0.0001 * rplcFamSymbol_cntr.Z);
            rplc_famInst.Location.Move(newPlace_diff_bk);
        }

        #endregion    // Function: Nudge Fitting to make force revit to reacttch the pipe properly... currently there is missalignments and gaps.


        #region  //  Set Radius or Diameter of rplc_famInst Fitting Based on the associated Pipe or Flange connector provided.
        public void resizeFittingConnPipeFlange(Connector famInst_opPipeConn_FittFacing, Connector rplc_fam_conn)
        {
            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in resizeFittingConnector. \n.Change Size for Replacement Family Instance:\n Get Pipe Element\n@@@@@@@@@@@@@@@@@@@@@");
            Document actvDoc = rplc_fam_conn.Owner.Document;
            Parameter fittingParameter;

            // Get the rplc_famInst element from the document using its instance ID
            Element rplc_famInst_el = actvDoc.GetElement(rplc_fam_conn.Owner.Id);
            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in resizeFittingConnPipeFlange. \n. In chngSize_rplc_famInst_trx \n rplc_famInst_el.Id: " + rplc_famInst_el.Id.ToString());

            if (actvDoc.GetElement(famInst_opPipeConn_FittFacing.Owner.Id) is Pipe)
            {    // Get the Pipe associated with this connector
                Pipe p = famInst_opPipeConn_FittFacing.Owner as Pipe;
                //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in resizeFittingConnPipeFlange. \n. In chngSize_rplc_famInst_trx \n Pipe ID: " + p.Id + " \n Pipe Diameter: " + p.Diameter);
                fittingParameter = GetAssocParam(rplc_fam_conn);
                using (Transaction resizeFitt2Pipe_trx = new Transaction(actvDoc, "Resize the Replacement Family Instance to Pipe; \n Set appropriate diameter to fitting port"))
                {
                    resizeFitt2Pipe_trx.Start();
                    fittingParameter.Set(p.Diameter / 2);   // In Pipe connectors, we need the radius not the  diameter... so devide by two.  In Flange mode, then the parameter ID takes care of whether radius or diameter is used.
                    resizeFitt2Pipe_trx.Commit();
                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "Finished Trx: reszFitt2Pipe_trx: ");
                }
                //fittingParameter.Set(p.Diameter);
                //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in resizeFittingConnPipeFlange. \n. In chngSize_rplc_famInst_trx \n \n rplc_famInst Diameter: " + fittingParameter.AsDouble());
            }
            else
            {
                using (Transaction resizeFitt2Flange_trx = new Transaction(actvDoc, "Resize the Replacement Family Instance to Flange; \n Set appropriate diameter to fitting port"))
                {
                    resizeFitt2Flange_trx.Start();
                    Parameter flangeParam;
                    flangeParam = GetAssocParam(famInst_opPipeConn_FittFacing);
                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in resizeFittingConnPipeFlange. \n. In chngSize_rplc_famInst_trx \n Flange ID: " + famInst_opPipeConn_FittFacing.Owner.Id + " \n\n Flange Diameter/RAD: " + flangeParam.AsDouble());
                    fittingParameter = GetAssocParam(rplc_fam_conn);
                    fittingParameter.Set(flangeParam.AsDouble());
                    resizeFitt2Flange_trx.Commit();
                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "Finished Trx: resizeFitt2Flange_trx: ");
                }
            }
        }
        #endregion  // Set Radius or Diameter of rplc_famInst Fitting Based on the flanges.


        #region  //  retrieves the parameter that holds the Diam or Radius of that connector (depends which was used)
        public Parameter GetAssocParam(Connector rplc_fam_conn)
        {
            Document actvDoc = rplc_fam_conn.Owner.Document;
            MEPFamilyConnectorInfo connectorInfo;
            ElementId parameterId;
            ParameterElement parameterElement;
            Parameter fittingParameter = null;

            // Get the rplc_famInst element from the document using its instance ID
            Element rplc_famInst_el = actvDoc.GetElement(rplc_fam_conn.Owner.Id);
            // Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in GetAssocParam. \n. In chngSize_rplc_famInst_trx \n rplc_famInst_el.Id: " + rplc_famInst_el.Id.ToString());

            // Find the associated parameter that holds the Diameter/Radius information:
            //// - get connector's MEPFamilyConnectorInfo
            connectorInfo = (MEPFamilyConnectorInfo)rplc_fam_conn.GetMEPConnectorInfo();
            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in GetAssocParam. \n. In chngSize_rplc_famInst_trx \n \n connectorInfo: " + connectorInfo.ToString());

            //// - get the builtin diam/radius parameter id from te MEPFamiltyConnectorInfo
            parameterId = connectorInfo.GetAssociateFamilyParameterId(new ElementId(BuiltInParameter.CONNECTOR_DIAMETER));
            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in GetAssocParam. \n. In chngSize_rplc_famInst_trx \n\n parameterId: " + parameterId);

            //// Here we have to check if the parameter ID for Diam  is null in which case we have to use that of the Radius. 
            if (parameterId != ElementId.InvalidElementId)  // If Diameter Parameter DOES exist then use it.
            {
                //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in GetAssocParam. \n. In chngSize_rplc_famInst_trx \n \n Diameter parameterId is NOT NULL  !!!!!!");
                // - get the parameterElement from that id in the active document.
                parameterElement = (ParameterElement)actvDoc.GetElement(parameterId);
                // - use this element definition of that parameter.
                fittingParameter = rplc_fam_conn.Owner.get_Parameter(parameterElement.GetDefinition());
                // - use this definition to access its value in the fitting.
                //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in GetAssocParam. \n fitting Diameter associated value: " + fittingParameter.GetType());

                // set that associated parameter value to the pipe.Diameter/Radius
                return fittingParameter;

            }
            else  // If Daiameter parameter does NOT exist then check radius parameter.
            {
                // - get the builtin diam/radius parameter id from te MEPFamiltyConnectorInfo
                parameterId = connectorInfo.GetAssociateFamilyParameterId(new ElementId(BuiltInParameter.CONNECTOR_RADIUS));
                // - get the parameterElement from that id in the active document.


                if (parameterId != ElementId.InvalidElementId)
                {
                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in resizeFittingConnector. \n. In chngSize_rplc_famInst_trx \n Raduis parameterId is NOT null");

                    parameterElement = (ParameterElement)actvDoc.GetElement(parameterId);
                    // - use this element definition of that parameter.
                    fittingParameter = rplc_fam_conn.Owner.get_Parameter(parameterElement.GetDefinition());
                    // - use this definition to access its value in the fitting.
                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in GetAssocParam. \n fitting Diameter associated value: " + fittingParameter.GetType());

                    // set that associated parameter value to the pipe.Diameter/Radius
                    return fittingParameter;
                }  // if Diam is null
                else    // If neither Diam and Radius parameters exist, then exit with Dialog warning.
                {
                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", " I am in GetAssocParam. \n. Warning \n Could not Find a Diameter or Radius Parameter for setting the Fitting Size.");
                }  // else Both Diam and radius null
            } // else  if Diam null
              //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in GetAssocParam. \n. Finish Trx: \n Set Radius or Diameter of rplc_famInst Fitting Based on the flanges.");
            return fittingParameter;
        }
        #endregion    //  retrieves the parameter that holds the Diam or Radius of that connector (depends which was used)



        #region  //Get the opposing connector, for example the connector on the pipe end that is attached to the fitting on that pipe's end.
        // there is one glitch here to keep an eye, on.  if the opposing connector is a flange, then we need to get the opposing connector is not that of a pipe.
        // so we either have to accept that and not make a priority, or get the check for it and return the connector of the pipe that is attached to that flange on the other side of the fitting.
        // We see this in pipe types that require a Band connection (a flange) such ast Cast Iron No-Hub.  with this pipe type the fittings are joined by a band(flange) and that is ID that the current 
        // code is returning for the  opposing connector insstea of htat of the pipe associated with that Band(flange)
        static public Connector getOpposingConnector(UIDocument uiDoc, Connector currConn, int excPartType)
        {
            Document actvDoc = uiDoc.Document as Document;
            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingConnector.\n Get Active Doc:  actvDoc !!!  ");
            double erMargin = 0.001;
            try
            {
                //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingConnector. currConn.AllRefs Count:  " + currConn.AllRefs.Cast<Connector>().ToList().Count);
                foreach (Connector ar_conn in currConn.AllRefs)
                {
                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingConnector. \n In AllRefs Loop");
                    if (((ar_conn.Origin.X - currConn.Origin.X) / currConn.Origin.X < erMargin) && ((ar_conn.Origin.Y - currConn.Origin.Y) / currConn.Origin.Y < erMargin) && ((ar_conn.Origin.Z - currConn.Origin.Z) / currConn.Origin.Z < erMargin))
                    {
                        //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingConnector. Found opposing Connector.  Info:  \n" + currConn.Origin.ToString() + "\n     vs \n" + ar_conn.Origin.ToString());
                        Element currPipe_el = actvDoc.GetElement(ar_conn.Owner.Id);
                        FamilyInstance famInst = currPipe_el as FamilyInstance;

                        //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingConnector. \n Found the FamilyInstance");
                        Parameter partTypeParam = famInst.Symbol.Family.get_Parameter(BuiltInParameter.FAMILY_CONTENT_PART_TYPE);

                        if (partTypeParam != null)
                        {
                            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingConnector. \n PartTypeParam is NOT  null !!");
                            PartType partType = (PartType)partTypeParam.AsInteger(); // e.g. 13 = Connection, 7 = Transition

                            if ((int)partType == excPartType)
                            {

                                //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "\n Opposing  Connector is of the excluded Family  !!!, \n Recurse !!!!!!!");
                                //recurse
                                getOpposingConnector(uiDoc, ar_conn, excPartType);
                            }
                            else
                            {
                                //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "\n Opposing  Connector is what we are looking for::::\n" + partType);
                                return ar_conn;
                            }
                        }

                        return null;
                    }
                    else
                    {
                        //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingConnector. Did not Find opposing Connector.  Info:  \n" + currConn.Origin.ToString() + "\n   vs \n" + ar_conn.Origin.ToString());
                        return null;
                    }
                }
            }
            catch (System.Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingConnector. Exception:  \n" + ex);
            }

            return null;
        }
        #endregion  //Get the opposing connector, for example the connector on the pipe end that is attached to the fitting on that pipe's end.

        #region  //  Function: Get the opposing Connectors.
        static public Connector getOpposingPipeConnector(Connector currConn)
        {
            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector.");

            Document actvDoc = currConn.Owner.Document;
            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector.\n Get Active Doc:  actvDoc !!!  ");
            double erMargin = 0.001;
            Element currConn_el = actvDoc.GetElement(currConn.Owner.Id);


            string currConn_type = "";
            try
            {
                if (currConn_el is Pipe) { currConn_type = "Pipe"; }
                //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector.\n\n currConn.AllRefs Count:  " + currConn.AllRefs.Cast<Connector>().ToList().Count + "\n\n currConn Type is: " + currConn_type);
                foreach (Connector ar_conn in currConn.AllRefs)
                {
                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector. \n In AllRefs Loop");



                    if (ar_conn != null)
                    {
                        // look for physical connections
                        if (ar_conn.ConnectorType == ConnectorType.End ||
                            ar_conn.ConnectorType == ConnectorType.Curve ||
                            ar_conn.ConnectorType == ConnectorType.Physical)
                        {

                            if (((ar_conn.Origin.X - currConn.Origin.X) / currConn.Origin.X < erMargin) && ((ar_conn.Origin.Y - currConn.Origin.Y) / currConn.Origin.Y < erMargin) && ((ar_conn.Origin.Z - currConn.Origin.Z) / currConn.Origin.Z < erMargin))
                            {
                                //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector. \n ar_conn.Owner.Id: " + ar_conn.Owner.Id);
                                return ar_conn;
                            }
                            else
                            {
                                return null;
                            }
                        }
                    }





                }

            }
            catch (System.Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector. Exception:  \n" + ex);
            }

            return null;
        }
        #endregion  //Get the opposing Connectors.

        #region  //Get the opposing Pipe connector, using recursive calls.
        static public Connector getOpposingPipeConnector_recur(UIDocument uiDoc, Connector currConn, int recnum)
        {
            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector_recur.\n Recurse num: " + recnum);
            int recurseLimit = 3;
            if (recnum == recurseLimit)
            {
                Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector_recur.\n Recurse Limit Exceeded !!!  \n " + recnum);
                return null;
            }
            Document actvDoc = uiDoc.Document as Document;
            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector_recur.\n Get Active Doc:  actvDoc !!!  ");
            double erMargin = 0.001;
            try
            {
                //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector_recur. currConn.AllRefs Count:  " + currConn.AllRefs.Cast<Connector>().ToList().Count);
                foreach (Connector ar_conn in currConn.AllRefs)
                {
                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector_recur. \n In AllRefs Loop");
                    if (((ar_conn.Origin.X - currConn.Origin.X) / currConn.Origin.X < erMargin) && ((ar_conn.Origin.Y - currConn.Origin.Y) / currConn.Origin.Y < erMargin) && ((ar_conn.Origin.Z - currConn.Origin.Z) / currConn.Origin.Z < erMargin))
                    {
                        //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector_recur. Found opposing Connector.  Info:  \n" + currConn.Origin.ToString() + "\n     vs \n" + ar_conn.Origin.ToString());
                        Element currPipe_el = actvDoc.GetElement(ar_conn.Owner.Id);
                        //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector_recur. \n just got current Connector Element !!! \n");
                        if (currPipe_el is Pipe)
                        {
                            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector_recur \n Opposing  Connector is what we are looking for::::\n  Pipe !!");
                            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector_recur \n Opposing  Connector ID is: \n " + ar_conn.Owner.Id);
                            return ar_conn;
                        }
                        else
                        {
                            // get the owner fo the non-pipe element
                            FamilyInstance ar_conn_famInst = actvDoc.GetElement(ar_conn.Owner.Id) as FamilyInstance;
                            // get its list of connectors
                            List<Connector> ar_conn_famInst_connectors = ar_conn_famInst.MEPModel.ConnectorManager.Connectors.Cast<Connector>().ToList();
                            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector_recur\n Opposing  Connector \n number of connectors:" + ar_conn_famInst_connectors.Count);

                            // get the connector with the opposing BasisZ with current connector ar_conn
                            foreach (Connector acfc in ar_conn_famInst_connectors)
                            {

                                // Do they have Opposing Z directions .. or is angle between connector equal to 180
                                double currAngle = acfc.CoordinateSystem.BasisZ.AngleTo(ar_conn.CoordinateSystem.BasisZ);

                                if (System.Math.Abs(currAngle - System.Math.PI) < 0.05 * System.Math.PI)
                                {
                                    //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector_recur \n Opposing  Connector is of the excluded Family  !!!, \n Recurse !!!!!!!");
                                    //recurse
                                    //// it is important to add return here, otherwise, the recursed call will come back here and the original function call will result in returning a null per code below.
                                    return getOpposingPipeConnector_recur(uiDoc, acfc, recnum + 1);

                                }

                            }
                            return null;
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingPipeConnector_recur. Exception:  \n" + ex);
            }

            return null;
        }
        #endregion  //Get the opposing Pipe connector, using recursive calls.

        #region  // Find what a given pipe connector go to what rplacement fitting instance (rplc_famInst) connector based on angle between connector and original fitting instance (famInst)
        public Connector connectTwoConn_basedOnAngleOnPrevFitting(Connector OpposingConn, FamilyInstance rplc_famInst)
        {
            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in connectTwoConn_basedOnAngleOnPrevFitting:  ");
            // get rplc_famInst connectors
            List<Connector> rplc_famInst_connectors = rplc_famInst.MEPModel.ConnectorManager.Connectors.Cast<Connector>().ToList();
            int i = 0;
            foreach (Connector rplc_conn in rplc_famInst_connectors)
            {
                double currAngle;

                // Do they have Opposing Z directions .. or is angle between connector equal to 180
                currAngle = rplc_conn.CoordinateSystem.BasisZ.AngleTo(OpposingConn.CoordinateSystem.BasisZ);
                //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in connectTwoConn_basedOnAngleOnPrevFitting:\n  " + "Angle = " + (currAngle * 180 / System.Math.PI).ToString() + "\n\n OpposingConn ID = " + OpposingConn.Owner.Id);
                if (System.Math.Abs(currAngle - System.Math.PI) < 0.05 * System.Math.PI)
                {
                    try
                    {
                        rplc_conn.ConnectTo(OpposingConn);
                        //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in connectTwoConn_basedOnAngleOnPrevFitting: \n Finished Connecting at " + "Angle = " + (currAngle * 180 / System.Math.PI).ToString() + "\n i = " + i);
                        i++;
                        return rplc_conn;
                    }
                    catch (System.Exception ex)
                    {
                        Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in connectTwoConn_basedOnAngleOnPrevFitting. \n\n Exception: " + ex);
                        return null;
                    }
                }
            }
            //Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in connectTwoConn_basedOnAngleOnPrevFitting: \n\n Exiting!! ");
            return null;
        }
        #endregion  // Find what pipe connectors go to what rplacement fitting instance (rplc_famInst) connector based on an original fitting instance (famInst)

        #region // Use the Transform of a Family and Assign it to another
        //https://stackoverflow.com/questions/52428481/given-two-family-instances-with-the-same-locationpoint-how-can-i-get-instance-1
        public System.Tuple<double, XYZ> GetAxisAngleFromMatrix(double[][] m)
        {
            double angleOfRotation;
            XYZ axisOfRotation;

            double angle, x, y, z; // variables for result
            double epsilon = 0.01; // margin to allow for rounding errors
            double epsilon2 = 0.1; // margin to distinguish between 0 and 180 degrees
                                   // optional check that input is pure rotation, 'isRotationMatrix' is defined at:
                                   // https://www.euclideanspace.com/Maths/algebra/matrix/orthogonal/rotation/

            if ((System.Math.Abs(m[0][1] - m[1][0]) < epsilon)
              && (System.Math.Abs(m[0][2] - m[2][0]) < epsilon)
              && (System.Math.Abs(m[1][2] - m[2][1]) < epsilon))
            {
                // singularity found
                // first check for identity matrix which must have +1 for all terms
                //  in leading diagonaland zero in other terms
                if ((System.Math.Abs(m[0][1] + m[1][0]) < epsilon2)
                  && (System.Math.Abs(m[0][2] + m[2][0]) < epsilon2)
                  && (System.Math.Abs(m[1][2] + m[2][1]) < epsilon2)
                  && (System.Math.Abs(m[0][0] + m[1][1] + m[2][2] - 3) < epsilon2))
                {
                    // this singularity is identity matrix so angle = 0
                    angleOfRotation = 0;
                    axisOfRotation = new XYZ(1, 0, 0);

                    return new System.Tuple<double, XYZ>(angleOfRotation, axisOfRotation);
                }

                // otherwise this singularity is angle = 180
                angle = System.Math.PI;
                double xx = (m[0][0] + 1) / 2;
                double yy = (m[1][1] + 1) / 2;
                double zz = (m[2][2] + 1) / 2;
                double xy = (m[0][1] + m[1][0]) / 4;
                double xz = (m[0][2] + m[2][0]) / 4;
                double yz = (m[1][2] + m[2][1]) / 4;
                if ((xx > yy) && (xx > zz))
                { // m[0][0] is the largest diagonal term
                    if (xx < epsilon)
                    {
                        x = 0;
                        y = 0.7071;
                        z = 0.7071;
                    }
                    else
                    {
                        x = System.Math.Sqrt(xx);
                        y = xy / x;
                        z = xz / x;
                    }
                }
                else if (yy > zz)
                { // m[1][1] is the largest diagonal term
                    if (yy < epsilon)
                    {
                        x = 0.7071;
                        y = 0;
                        z = 0.7071;
                    }
                    else
                    {
                        y = System.Math.Sqrt(yy);
                        x = xy / y;
                        z = yz / y;
                    }
                }
                else
                { // m[2][2] is the largest diagonal term so base result on this
                    if (zz < epsilon)
                    {
                        x = 0.7071;
                        y = 0.7071;
                        z = 0;
                    }
                    else
                    {
                        z = System.Math.Sqrt(zz);
                        x = xz / z;
                        y = yz / z;
                    }
                }

                angleOfRotation = angle;
                axisOfRotation = new XYZ(x, y, z); // return 180 deg rotation

                return new System.Tuple<double, XYZ>(angleOfRotation, axisOfRotation);
            }
            // as we have reached here there are no singularities so we can handle normally
            double s = System.Math.Sqrt((m[2][1] - m[1][2]) * (m[2][1] - m[1][2])
              + (m[0][2] - m[2][0]) * (m[0][2] - m[2][0])
              + (m[1][0] - m[0][1]) * (m[1][0] - m[0][1])); // used to normalise
            if (System.Math.Abs(s) < 0.001) s = 1;
            // prevent divide by zero, should not happen if matrix is orthogonal and should be
            // caught by singularity test above, but I've left it in just in case
            angle = System.Math.Acos((m[0][0] + m[1][1] + m[2][2] - 1) / 2);
            x = (m[2][1] - m[1][2]) / s;
            y = (m[0][2] - m[2][0]) / s;
            z = (m[1][0] - m[0][1]) / s;

            angleOfRotation = angle;
            axisOfRotation = new XYZ(x, y, z);
            return new System.Tuple<double, XYZ>(angleOfRotation, axisOfRotation);
        }
        #endregion // Use the Transform of a Family and Assign it to another

    }  // end of Class rplcFamilyInstances_CN

}  // namespace KDS_Module







