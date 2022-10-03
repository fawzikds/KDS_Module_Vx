using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace KDS_Module
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public partial class WyeToCombo : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DialogBox(commandData.Application.ActiveUIDocument);
            return Result.Succeeded;
        }

        #region // DialogBox Function to prompt user with which tool they want to choose, and then execution of specified command //
        public void DialogBox(UIDocument uidoc)
        {
            TaskDialog mainDialog = new TaskDialog("Change Pipe Fitting")
            {
                MainInstruction = "Please choose an option",
                MainContent = "Which command would you like to run?",
                CommonButtons = TaskDialogCommonButtons.Close,
                DefaultButton = TaskDialogResult.Close,
                FooterText = "KDS Plumbing and Heating Services",
                TitleAutoPrefix = false,
                MainIcon = TaskDialogIcon.TaskDialogIconInformation,
            };
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Replace Wye With Combo");
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Replace Wye With Santee");

            TaskDialogResult tResult = mainDialog.Show();

            bool Yes = true;

            if (TaskDialogResult.CommandLink1 == tResult)
            { Yes = true; }
            else if (TaskDialogResult.CommandLink2 == tResult)
            { Yes = false; }
            else
            { return; }

            if (Yes)
            {
                ChangeFitting(uidoc.Application); // end if (Yes)
            }

            if (!Yes)
            {
                TaskDialog.Show("WyeToSantee", "No command yet.");

            }
        } // end kdsMacro
        private static IEnumerable<Type> GetTypesSafely(Assembly assembly)
        {
            try
            {
                return assembly.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                return ex.Types.Where(x => x != null);
            }
        }
        #endregion

        public void ChangeFitting(UIApplication a)
        {
            UIDocument uiDoc = a.ActiveUIDocument;
            Document actvDoc = uiDoc.Document;

            #region  // Define Family Names for collector search.

            // It looks like for KDS_Char_CI_NH, i need to use the Wye as default for designer, and then later replace with Combos.
            // Since using San Tee as Default adds a 45 degrees to make a wierd fitting combination for a Wye.

            string selectString = "KDS_Char_Ci_NH_Wye"; // NAME OF OBJECT TO BE REPLACED //
            const string param_name = "Angle"; // NAME OF PARAMETER BEING COMPARED / ANALYZED //
            string desiredAngleValue = "90.00°"; // PARAMETER VALUE //
            string replaceString = "KDS_Char_CI_NH_Combo"; // NAME OF REPLACING OBJECT 
            #endregion  // Define Family Names for collector search.

            #region // Get current family to be replaced based on its name

            FilteredElementCollector currCollector = null;
            currCollector = new FilteredElementCollector(actvDoc);
            Family currFam = null;
            currFam = currCollector.OfClass(typeof(Family)).OfType<Family>().FirstOrDefault(f => f.Name.Equals(selectString));

            FamilySymbol currFamSymbol = null;
            currFamSymbol = actvDoc.GetElement(currFam.GetFamilySymbolIds().FirstOrDefault()) as FamilySymbol;
            currFamSymbol.Activate();
            #endregion

            #region // Get the replacement Family based on its name.

            FilteredElementCollector rplcCollector = new FilteredElementCollector(actvDoc);

            Family rplcFam = rplcCollector.OfClass(typeof(Family)).OfType<Family>().FirstOrDefault(f => f.Name.Equals(replaceString));

            FamilySymbol rplcFamSymbol = actvDoc.GetElement(rplcFam.GetFamilySymbolIds().FirstOrDefault()) as FamilySymbol;
            rplcFamSymbol.Activate();

            FamilyInstance rplc_famInst;

            #endregion //Get the replacement Family  based on its name. 

            #region // Get All Family instances to be replaced in Project 

            FilteredElementCollector instCollector = new FilteredElementCollector(actvDoc);

            List<FamilyInstance> allFamInst = instCollector.OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                                                                                           .Where(inst => inst.Symbol.Id.Equals(currFamSymbol.Id) && inst.LookupParameter(param_name)
                                                                                           .AsValueString() == desiredAngleValue).ToList();
            List<FamilyInstance> filteredFamInst = new List<FamilyInstance>();

            bool Yes = false;

            #region // Filtering based on orientation //

            foreach (FamilyInstance inst in allFamInst)
            {
                if (inst.FacingOrientation.ToString().Equals(new XYZ(0, 0, 1).Normalize().ToString()))
                {
                    if (inst.HandOrientation.ToString().Equals(new XYZ(1, 0, 0).Normalize().ToString()) || inst.HandOrientation.ToString().Equals(new XYZ(-1, 0, 0).Normalize().ToString()))
                    {
                        filteredFamInst.Add(inst);
                        Yes = true;
                    }

                    if (inst.HandOrientation.ToString().Equals(new XYZ(0, 1, 0).Normalize().ToString()) || inst.HandOrientation.ToString().Equals(new XYZ(0, -1, 0).Normalize().ToString()))
                    {
                        filteredFamInst.Add(inst);
                        Yes = true;
                    }
                }
                if (inst.FacingOrientation.ToString().Equals(new XYZ(0, 0, -1).Normalize().ToString()))
                {
                    if (inst.HandOrientation.ToString().Equals(new XYZ(0, -1, 0).Normalize().ToString()) || inst.HandOrientation.ToString().Equals(new XYZ(0, 1, 0).Normalize().ToString()))
                    {
                        filteredFamInst.Add(inst);
                        Yes = true;
                    }

                    if (inst.HandOrientation.ToString().Equals(new XYZ(-1, 0, 0).Normalize().ToString()) || inst.HandOrientation.ToString().Equals(new XYZ(1, 0, 0).Normalize().ToString()))
                    {
                        filteredFamInst.Add(inst);
                        Yes = true;
                    }

                    if (Yes == false)
                    {
                        TaskDialog.Show("temp", "current Fitting will become a combo");
                    }
                }
            }
            #endregion
              
            #endregion // Get All Family instances to be replaced in Project

            #region // ReplaceFamilies 

            foreach (FamilyInstance famInst in filteredFamInst)
            {
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

                    #region  // Get Connectors of famInst.... I need these to get the opposing connectors, i.e. those of the pipes connected to this fitting.

                    List<Connector> famInst_connectors = famInst.MEPModel.ConnectorManager.Connectors.Cast<Connector>().ToList();

                    Connector conn_0 = famInst_connectors[0];
                    Connector conn_1 = famInst_connectors[1];
                    Connector conn_2 = famInst_connectors[2];

                    if (famInst_connectors == null || famInst_connectors.Count == 0) { TaskDialog.Show("test", "famInst_connectors is null or empty"); }

                    #endregion  // Get Connectors of famInst.

                    //It seems that get opposing connector is returning the husky Band in this case and not the pipe //

                    #region  // Get opposing connectors to the family instance connectors.. i need these to make the connection when connecting the new fitting to the same pipes, later on

                    List<Connector> famInst_opConn = new List<Connector>
                    {
                        GetOpposingConnector(uiDoc, famInst_connectors[0]),
                        GetOpposingConnector(uiDoc, famInst_connectors[1]),
                        GetOpposingConnector(uiDoc, famInst_connectors[2])
                    };
                    if (famInst_opConn == null || famInst_opConn.Count == 0) { TaskDialog.Show("test", "famInst_opConn is null or empty"); }

                    // TaskDialog.Show("rplcFamilyInstances_CN", "Opposing Connector IDs:\n" + famInst_opConn[0].Owner.Id + "\n" + famInst_opConn[1].Owner.Id + "\n" + famInst_opConn[2].Owner.Id);

                    #endregion  // Get opposing connectors to the family instance connectors

                    #region // Insert replacement Family

                    // Get Center of Fitting //
                    XYZ rplcFamSymbol_center = ((LocationPoint)famInst?.Location)?.Point;

                    // Get the Level of famInst //
                    Level famInst_refLevel = actvDoc.GetElement(famInst.LevelId) as Level;

                    #region  // Delete Original Family Instance: famInst

                    using (Transaction delFamInst_trx = new Transaction(actvDoc, "Delete Original Family Instance: famInst"))
                    {
                        // Delete the current  Fitting... if i delete before i cannot get above conn_0.Owner.Id //
                        delFamInst_trx.Start();
                        actvDoc.Delete(famInst.Id);
                        delFamInst_trx.Commit();
                    }
                    #endregion //Delete Original Family Instance: famInst

                    #region // Replace Family Instance

                    using (Transaction rplc_famInst_trx = new Transaction(actvDoc, "Replace Family Instance"))
                    {
                        rplc_famInst_trx.Start();
                        rplc_famInst = actvDoc.Create.NewFamilyInstance(rplcFamSymbol_center, rplcFamSymbol, famInst_refLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                        string facingOrientation_txt = rplc_famInst.FacingOrientation.ToString();
                        rplc_famInst_trx.Commit();

                    }//end transaction rplc_famInst_trx
                    #endregion // Replace Family Instance

                    #region // Rotate rplcFam_inst based on Matrix
                    // for a wye fitting, i can find colinear pipes (pipes that would form a line)
                    // so, in the fitting find the connectors that form a line.
                    // find the angle between the two lines, rotate the fitting by that angle so they line up.
                    //checked third connector with the other 3rd connector. if same angle and opposite do nothing
                    // if same angle but same direction, then rotate 180 degrees
                    // if 90 degrees flip around an axis perpendicular to the colinear lines
                    // if 270 flip around the colinear axis.

                    //    but what about a cross.? Jake: Unsure what you mean by a cross

                    using (Transaction RotateRplcFamInst_trx = new Transaction(actvDoc, "Rotate rplcFam_inst based on Matrix"))
                    {
                        RotateRplcFamInst_trx.Start();
                        var anAx = GetAxisAngleFromMatrix(famInst_matrix);
                        double angleOfRotation = anAx.Item1;
                        XYZ axisOfRotation = anAx.Item2;

                        Line rotationLine = Line.CreateUnbound(famInst_transform.Origin, axisOfRotation);
                        ElementTransformUtils.RotateElement(actvDoc, rplc_famInst.Id, rotationLine, angleOfRotation);
                        RotateRplcFamInst_trx.Commit();
                    }
                    #endregion // Rotate rplcFam_inst based on Matrix

                    #region // Change Routing preferences to use The replacement fitting as Default.

                    using (Transaction changeRoutingPref_trx = new Transaction(actvDoc, "Rotate rplcFam_inst based on Matrix"))
                    {
                        changeRoutingPref_trx.Start();
                        Connector currPipeConn = famInst_opConn[0];
                        if (currPipeConn is Connector) { TaskDialog.Show("rplcFamilyInstances_CN", "I am in rplcFamilies_CN.execute.\n currPipeConn = Connector !!!  "); }
                        else { TaskDialog.Show("rplcFamilyInstances_CN", "I am in rplcFamilies_CN.execute.\n currPipeConn is NOT Connector !!!  "); }


                        //var allRefs = currPipeConn.AllRefs;
                        List<Element> ConnectedElements = new List<Element>();

                        foreach (Connector connector in famInst_connectors)
                        {
                            ConnectedElements.Add(connector.Owner);
                            TaskDialog.Show("test", "Pipe: " + connector.Owner.Name);
                        }

                        //Pipe currPipe = (Pipe)ConnectedElements[0];


                        // Item SelectionDialogResult seems somthing is wrong with passing the pipe to the changeRouting.. function... i am not sure what it is now, 
                        // Pipe currPipe = currPipeConn.Owner as Pipe;

                        Pipe currPipe = ConnectedElements.LastOrDefault() as Pipe;

                        TaskDialog.Show("test", "currPipeConn.Owner is: " + currPipe.Name); //Currently Duriron Clamp Joint//  

                        TaskDialog.Show("rplcFamilyInstances_CN", "Got currPipe as a Pipe");
                        if (currPipe == null) { TaskDialog.Show("error", "currPipe is null"); }

                        //highlightElementById(uiDoc, currPipe.Id);
                        //TaskDialog.Show("rplcFamilyInstances_CN", "Highlighted Element");

                        Element currPipe_el = actvDoc.GetElement(famInst_opConn[0].Owner.Id);
                        TaskDialog.Show("rplcFamilyInstances_CN", "Got currPipe_el as element");
                        if (currPipe_el == null) { TaskDialog.Show("error", "currPipe_el is null"); }

                        if (currPipe_el.Category.Id == new ElementId(BuiltInCategory.OST_PipeFitting)) { TaskDialog.Show("rplcFamilyInstances_CN", "I am in rplcFamilies_CN.execute.\n currPipe IS a  PIPE Fitting !!!  "); }
                        //if (currPipe is Pipe) { Autodesk.Revit.UI.TaskDialog.Show("rplcFamilyInstances_CN", "I am in rplcFamilies_CN.execute.\n currPipe IS a  PIPE !!!  "); }
                        else { TaskDialog.Show("rplcFamilyInstances_CN", "I am in rplcFamilies_CN.execute.\n currPipe is NOT PIPE Fitting!!!  "); }

                        TaskDialog.Show("Test", "currPipe: " + currPipe_el);

                        TaskDialog.Show("Test", "currPipe_el: " + currPipe);

                        ChangeRoutingPreferencesOfTee(currPipe, "Junctions", rplcFam);  // how do i up the a family as default.. how is the rule written...
                        changeRoutingPref_trx.Commit();
                    }
                    #endregion // Change Routing preferences to use The replacement fitting as Default.

                    #region // Connect the Replacement Family Instance to pipes and set appropriate diameter to fitting port

                    using (Transaction ConnectRplcFamInst_trx = new Transaction(actvDoc, "Connect the Replacement Family Instance to pipes and set appropriate diameter to fitting port"))
                    {
                        ConnectRplcFamInst_trx.Start();

                        ConnectPipesToFittingBasedOnPrevFitting(famInst_opConn, famInst_connectors, rplc_famInst);

                        TaskDialog.Show("rplcFamilyInstances_CN", "Connect the Replacement Family Instance to pipes and set appropriate diameter to fitting port");

                        ConnectRplcFamInst_trx.Commit();
                    }
                    #endregion // Connect the Replacement Family Instance to pipes and set appropriate diameter to fitting port
                    #endregion

                }//end try
                catch (Exception e)
                {
                    TaskDialog.Show("rplcFamilyInstances_CN", "I am in rplcFamilies_CN.execute. Exception:  " + e.ToString());
                }// end catch
            }  // Foreach Family to edit
            #endregion

        }// end of Execute
        #region  //Get the opposing connector, for example the connector on the pipe end that is attached to the fitting on that pipe's end.
        static public Connector GetOpposingConnector(UIDocument uidoc, Connector currConn)
        {
            try
            {
                foreach (Connector ar_conn in currConn.AllRefs)
                {
                    if ((ar_conn.Origin.X == currConn.Origin.X) && (ar_conn.Origin.Y == currConn.Origin.Y) && (ar_conn.Origin.Z == currConn.Origin.Z))
                    {

                    }
                    else
                    {

                        return ar_conn;
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("rplcFamilyInstances_CN", "I am in getOpposingConnector. Exception:  \n" + ex);
            }
            return null;
        }
        #endregion  //Get the opposing connector, for example the connector on the pipe end that is attached to the fitting on that pipe's end.

        #region  // Find what pipe connectors go to what rplacement fitting instance (rplc_famInst) connector based on an original fitting instance (famInst)
        // And while i am at it, Check the size of the pipe (associated Connector) and apply it to the Fitting.
        // This concept here is a bit tricky, since a fitting does not need to have the same pipe diameter(or radius) on all of its ports.
        // There are reducing fitting, such as a reducing Tee or Reducing Wye.  Which means that we need to know the size of th pipe connector associated with one of the fitting connectors
        // after which we can get the pipe connector size and set its value to the fitting connector.
        // Although i know the associated connectors,as i used the opposing connector concept from the previous fitting, i will not be able to make use of it as i do not know the name of the parameter
        // it is assigned to within the fitting.
        // As an example, Take a Tee familt that may allow reducing connections.  Such family will have at least 2 Diameters/radii to change. one for the main ports and one for the branches,
        // something like " Main Diameter" and "Vent Diameter".
        // IF i know the size of the pipe attached to one of my fitting branches, how can i know what it is called within the Fitting family (i.e. Vent diameter) so i can access it and change its value?
        // i get the size from the pipe (opposing) connector
        // then i get the associated connectorElement for the connecotr on the fitting that i am dealing with now
        // then i set that associated connectorElement size to the pipe size.
        // so the associated connectorElement is to look inside the fitting family ans see what a connector is associated to.
        public void ConnectPipesToFittingBasedOnPrevFitting(List<Connector> famInst_opConn_lst, List<Connector> famInst_connectors, FamilyInstance rplc_famInst)
        {
            TaskDialog.Show("rplcFamilyInstances_CN", "I am in connectPipesToFittingBasedOnPrevFitting:  ");

            // get rplc_famInst connectors
            List<Connector> rplc_famInst_connectors = rplc_famInst.MEPModel.ConnectorManager.Connectors.Cast<Connector>().ToList();

            // foreach connector in rplc_famInst_connectors loop thru famInst Connectors and find a matching Z direction then take the index of that to connect that connector to the opposing connector to famInst.
            // I cannot compare origins, since the inserted fitting may not be of the same size, and thus it will not have same origins, or even close enough origins.

            foreach (Connector rplc_conn in rplc_famInst_connectors)
            {
                double currAngle;
                int i = 0;
                foreach (Connector famInst_conn in famInst_opConn_lst)
                {
                    // Do they have Opposing Z directions .. or is angle between connector equal to 180
                    currAngle = rplc_conn.CoordinateSystem.BasisZ.AngleTo(famInst_conn.CoordinateSystem.BasisZ);

                    if (Math.Abs(currAngle - Math.PI) < 0.05 * Math.PI)
                    {
                        // get pipe size. // SAME ISSUE HERE WITH PIPE BEING NULL
                        Pipe p = famInst_opConn_lst[i].Owner as Pipe;
                        double rplc_conn_diam = p.Diameter;

                        // get associated connectorElement with the curr rplc_conn
                        // get connectorElement radius or diameter
                        // set fitting instance rplc_famInst associated parameter to the size of the pipe.
                        TaskDialog.Show("rplcFamilyInstances_CN", "I am in connectPipesToFittingBasedOnPrevFitting.  rplc_conn Radius:\n  " + rplc_conn.Radius * 12); // Why multiply by 12?
                        rplc_conn.Radius = rplc_conn_diam / 2;

                        rplc_conn.ConnectTo(famInst_opConn_lst[i]);
                        // TaskDialog.Show("rplcFamilyInstances_CN", "I am in connectPipesToFittingBasedOnPrevFitting:\n  " + "Angle = " + (currAngle * 180/System.Math.PI ).ToString() + "\n i = " + i);

                        break;
                    }
                    i++;
                }
            }
        }
        #endregion  // Find what pipe connectors go to what rplacement fitting instance (rplc_famInst) connector based on an original fitting instance (famInst)

        #region // Use the Transform of a Family and Assign it to another
        //https://stackoverflow.com/questions/52428481/given-two-family-instances-with-the-same-locationpoint-how-can-i-get-instance-1
        public Tuple<double, XYZ> GetAxisAngleFromMatrix(double[][] m)
        {
            double angleOfRotation;
            XYZ axisOfRotation;
            double angle, x, y, z; // variables for result
            double epsilon = 0.01; // margin to allow for rounding errors
            double epsilon2 = 0.1; // margin to distinguish between 0 and 180 degrees
                                   // optional check that input is pure rotation, 'isRotationMatrix' is defined at:
                                   // https://www.euclideanspace.com/Maths/algebra/matrix/orthogonal/rotation/

            if ((Math.Abs(m[0][1] - m[1][0]) < epsilon) && (Math.Abs(m[0][2] - m[2][0]) < epsilon) && (Math.Abs(m[1][2] - m[2][1]) < epsilon))
            {
                // singularity found
                // first check for identity matrix which must have +1 for all terms in leading diagonal and zero in other terms
                if ((Math.Abs(m[0][1] + m[1][0]) < epsilon2) && (Math.Abs(m[0][2] + m[2][0]) < epsilon2) && (Math.Abs(m[1][2] + m[2][1]) < epsilon2) && (Math.Abs(m[0][0] + m[1][1] + m[2][2] - 3) < epsilon2))
                {
                    // this singularity is identity matrix so angle = 0
                    angleOfRotation = 0;
                    axisOfRotation = new XYZ(1, 0, 0);

                    return new Tuple<double, XYZ>(angleOfRotation, axisOfRotation);
                }
                // otherwise this singularity is angle = 180
                angle = Math.PI;
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
                        x = Math.Sqrt(xx);
                        y = xy / x;
                        z = xz / x;
                    }
                }
                else if ((yy > zz) && (yy > xx))
                { // m[1][1] is the largest diagonal term
                    if (yy < epsilon)
                    {
                        x = 0.7071;
                        y = 0;
                        z = 0.7071;
                    }
                    else
                    {
                        y = Math.Sqrt(yy);
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
                        z = Math.Sqrt(zz);
                        x = xz / z;
                        y = yz / z;
                    }
                }
                angleOfRotation = angle;
                axisOfRotation = new XYZ(x, y, z); // return 180 deg rotation

                return new Tuple<double, XYZ>(angleOfRotation, axisOfRotation);
            }
            // as we have reached here, there are no singularities so we can handle normally
            double s = Math.Sqrt((m[2][1] - m[1][2]) * (m[2][1] - m[1][2])
              + (m[0][2] - m[2][0]) * (m[0][2] - m[2][0])
              + (m[1][0] - m[0][1]) * (m[1][0] - m[0][1])); // used to normalise

            if (Math.Abs(s) < 0.001) s = 1;
            // prevent divide by zero, should not happen if matrix is orthogonal and should be
            // caught by singularity test above, but I've left it in just in case

            angle = Math.Acos((m[0][0] + m[1][1] + m[2][2] - 1) / 2);
            x = (m[2][1] - m[1][2]) / s;
            y = (m[0][2] - m[2][0]) / s;
            z = (m[1][0] - m[0][1]) / s;

            angleOfRotation = angle;
            axisOfRotation = new XYZ(x, y, z);
            return new Tuple<double, XYZ>(angleOfRotation, axisOfRotation);
        }
        #endregion // Use the Transform of a Family and Assign it to another

        #region // Set Routing preferences on the fly.  i need this since replacing a Tee with a wye will break the connection once i set the size of the port diameter,

        // since revit sudenly wants to insert my default choice of a Tee.
        //https://forums.autodesk.com/t5/revit-api-forum/routing-preference-with-api-no-ui/td-p/9355761
        public void ChangeRoutingPreferencesOfTee(Pipe p, string grpTypStr, Family defaultFam)
        {
            TaskDialog.Show("rplcFamilyInstances_CN", "I am in changeRoutingPreferencesOfTee.");

            RoutingPreferenceRuleGroupType groupType = RoutingPreferenceRuleGroupType.Segments;
            if (grpTypStr == "Segments") { groupType = RoutingPreferenceRuleGroupType.Segments; }        //	The segment types. (e.g. pipe stocks)
            if (grpTypStr == "Elbows") { groupType = RoutingPreferenceRuleGroupType.Elbows; }         //	The elbow types.  (e.g. 90, 45, quarter Bends, Eighth Bend, Short Sweeps..etc.)
            if (grpTypStr == "Junctions") { groupType = RoutingPreferenceRuleGroupType.Junctions; }     //	The junction types(e.g.takeoff, tee, wye,, Combos, tap).
            if (grpTypStr == "Crosses") { groupType = RoutingPreferenceRuleGroupType.Crosses; }         //	The cross types ( Double Combos, Double Wyes, Double San Tee, Cross).
            if (grpTypStr == "Transitions") { groupType = RoutingPreferenceRuleGroupType.Transitions; } //	The transition types(Note that the multi-shape transitions may have their own groups).
            if (grpTypStr == "Unions") { groupType = RoutingPreferenceRuleGroupType.Unions; }           //	The segment types (Unions)

            TaskDialog.Show("rplcFamilyInstances_CN", "I am in changeRoutingPreferencesOfTee.\n  Finished assigning value to groupType");

            // Default, means Rule Index is 0.
            int ruleIndex = 0;
            TaskDialog.Show("rplcFamilyInstances_CN", "I am in changeRoutingPreferencesOfTee.\n  Finished assigning value to ruleIndex");

            Document actvDoc = p.Document;

            RoutingPreferenceManager routePrefManager = p.PipeType.RoutingPreferenceManager;
            routePrefManager.PreferredJunctionType = PreferredJunctionType.Tee;
            TaskDialog.Show("rplcFamilyInstances_CN", "I am in changeRoutingPreferencesOfTee.\n  Finished assigning value to PreferredJunctionType");

            int initRuleCount = routePrefManager.GetNumberOfRules(RoutingPreferenceRuleGroupType.Junctions);

            #region // Setting a new default, as i recall, does not need to remove a previous rule. Assigning rule index to zero, will shift others down the list.
            /*
            for (int i = 0; i != initRuleCount; ++i)
            {
                routePrefManager.RemoveRule(RoutingPreferenceRuleGroupType.Junctions, 0);
            }
            */
            #endregion

            #region   // i Have the family loaded already.
            /*
            //Family tapFam = null;
            //string path =
            //    @"C:\ProgramData\Autodesk\RVT 2020\Libraries\US Imperial\Duct\Fittings\Round\Taps\Round Takeoff.rfa";
            //actvDoc.LoadFamily(path, out tapFam);
            */
            #endregion

            FamilySymbol symbol = null;
            ISet<ElementId> familySymbolIds = defaultFam.GetFamilySymbolIds();
            ElementId id = familySymbolIds.ElementAt(0);

            symbol = defaultFam.Document.GetElement(id) as FamilySymbol;
            TaskDialog.Show("rplcFamilyInstances_CN", "I am in changeRoutingPreferencesOfTee.\n  Finished assigning value to symbol");

            if ((!symbol.IsActive) && (symbol != null))
            {
                symbol.Activate();
                actvDoc.Regenerate();
            }

            RoutingPreferenceRule newRule = new RoutingPreferenceRule(symbol.Id, "Set Default Family");
            TaskDialog.Show("rplcFamilyInstances_CN", "I am in changeRoutingPreferencesOfTee.\n  Finished assigning value to newRule");
            using (Transaction transaction = new Transaction(actvDoc, "Routing Preference"))
            {
                transaction.Start();
                try
                {
                    routePrefManager.AddRule(groupType, newRule, ruleIndex);
                }
                catch (Exception ruleExcp)
                {
                    TaskDialog.Show("rplcFamilyInstances_CN", "I am in changeRoutingPreferencesOfTee. Exception in AddRule = " + ruleExcp);
                }
                transaction.Commit();
            }
        }
        #endregion
    }
    #region // Notes To Self on how to run WyeToCombo Command
    /*
    * CURRENT ISSUE TO WORK ON:
    
    * When replacing occurs, new fittings have 2 issues: the diameter is 8", should be 4". Also, opposite connector does not connect, only 1 of the connectors connects
    * CURRPIPE is null, which is why we get error in ChangeRoutingPreferenceOnTheFly Function. Specific error is p.document, says "p is null".
    * whne ChangeRoutingPreferenceOnTheFly is commented out, error still occurs because Pipe currPipe (p) is still used. To progress, need to find way to have pipe connected to fitting not be null.
     */
    #endregion
}
