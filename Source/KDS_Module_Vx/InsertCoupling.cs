

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;

using System;
using System.Collections.Generic;
using System.Linq;

namespace KDS_Module
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("4EF3D6AC-4B72-4C41-BC43-DF221ED58BCF")]
    public class InsertCoupling : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            SplitPipe(commandData.Application.ActiveUIDocument.Document);
            return Result.Succeeded;
        }

        #region // SplitPipe Function //
        public void SplitPipe(Document actvDoc)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            
            int i = 0;
            double couplingSpacing;
            double segmentLength;

            FilteredElementCollector pipeCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(Pipe));
            List<Pipe> pipes = new List<Pipe>();
            List<PipeTypeInfo> pipeType_list = new List<PipeTypeInfo>();

            LoadValues(pipeType_list);

            foreach (PipeTypeInfo pipeType in pipeType_list)
            {
                pipes = pipeCollector.Cast<Pipe>().Where(p => p.PipeType.Name.Equals(pipeType.Name) && p.LookupParameter("Length").AsDouble() > pipeType.ManufactureLength).ToList();

                if (pipes.Count > 0)
                {
                    foreach (Pipe p in pipes)
                    {
                        List<XYZ> insertionPoints = new List<XYZ>();
                        Curve pipeCurve = FindPipeCurve(p);

                        couplingSpacing = pipeType.Spacing_Dictionary[Math.Round(p.Diameter, 4)];

                        segmentLength = pipeType.ManufactureLength + 2 * couplingSpacing; // REMOVED (couplingSpacing / 2). Part of original dynamo.

                        //  Get a list of points to segment the pipe.

                        for (double length = pipeType.ManufactureLength + (couplingSpacing / 2); length < pipeCurve.Length; length = length + segmentLength)
                        {
                            XYZ temp = null;

                            temp = pipeCurve.Evaluate(pipeCurve.ComputeNormalizedParameter(length), true);

                            insertionPoints.Add(temp);

                        }

                        Connector connA = null;
                        Connector ConnB = null;
                        ElementId newPipeId;
                        Pipe newPipe;
                        List<Connector> newPipeConnectors = new List<Connector>();

                        foreach (XYZ point in insertionPoints)
                        {
                            using (Transaction breakPipe_trx = new Transaction(actvDoc, "break pipe Transaction"))
                            {
                                breakPipe_trx.Start();
                                newPipeId = PlumbingUtils.BreakCurve(actvDoc, p.Id, point);

                                newPipe = actvDoc.GetElement(newPipeId) as Pipe;

                                newPipeConnectors = newPipe.ConnectorManager.Connectors.Cast<Connector>().ToList();

                                breakPipe_trx.Commit();
                            }

                            foreach (Connector c in p.ConnectorManager.Connectors)
                            {
                                XYZ pc = c.Origin;
                                List<Connector> nearest = new List<Connector>();

                                foreach (Connector conn in newPipeConnectors)
                                {
                                    if (pc.DistanceTo(conn.Origin) < 0.01)
                                    {
                                        nearest.Add(conn);
                                    }
                                }

                                if (nearest.Count > 0)
                                {
                                    connA = c;
                                    ConnB = nearest[0];

                                    using (Transaction newUnionFitting_trx = new Transaction(actvDoc, "new Union Fitting Transaction"))
                                    {
                                        newUnionFitting_trx.Start();
                                        actvDoc.Create.NewUnionFitting(connA, ConnB);
                                        newUnionFitting_trx.Commit();
                                        i++;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            watch.Stop();
            var elapsedTime = watch.ElapsedMilliseconds;
            TaskDialog.Show("test", "# of couplings inserted: " + i + "\n execution time : " + elapsedTime / 1000.0 + " seconds");
        }
        #endregion

        #region // LoadValues Function //
        /*
         all Measurements in Dictionaries are in ft conversions. E.g: 3/16 inch = 0.0156 ft
         */
        static void LoadValues(List<PipeTypeInfo> AuthorList)
        {
            //Pipe Spacing: 3/16th inch (0.0156 ft)
            // Currently husky band is set at 1/4" (6.4mm) seperation between the 2 connectors... which is about 0.020833 ft.  so half of that is 0.01041667"
            Dictionary<double, double> CI_Spacing = new Dictionary<double, double>()
            {
                {0.1250, 1.0 / 48.0},
                {0.2083, 1.0 / 48.0},
                {0.2500, 1.0 / 48.0},
                {0.3333, 1.0 / 48.0},
                {0.4167, 1.0 / 48.0},
                {0.5000, 1.0 / 48.0},
                {0.6667, 1.0 / 48.0},
                {0.8333, 1.0 / 48.0},
                {1.0000, 1.0 / 48.0}
            };

            Dictionary<double, double> CUL_Spacing = new Dictionary<double, double>()
            {
                {0.0417, 0.0000},
                {0.0625, 0.0000},
                {0.0833, 0.0000},
                {0.1042, 0.0000},
                {0.1250, 0.0000},
                {0.2083, 0.0000},
                {0.2500, 0.0000},
                {0.3333, 0.0000},
                {0.4167, 0.0000},
                {0.5000, 0.0000},
                {0.6667, 0.0000},
                {0.8333, 0.0000},
                {1.0000, 0.0000}
            };

            Dictionary<double, double> BLK_GAL_Spacing = new Dictionary<double, double>()
            {
                {0.0417, 0.0000},
                {0.0625, 0.0000},
                {0.0833, 0.0000},
                {0.1042, 0.0000},
                {0.1250, 0.0000},
                {0.2083, 0.0000},
                {0.2500, 0.0000},
                {0.3333, 0.0000},
                {0.4167, 0.0000},
                {0.5000, 0.0000},
                {0.6667, 0.0000},
                {0.8333, 0.0000},
                {1.0000, 0.0000}
            };

            Dictionary<double, double> PVC_spacing = new Dictionary<double, double>()
            {
                {0.0417, 0.0156},
                {0.0625, 0.0156},
                {0.0833, 0.0156},
                {0.1042, 0.0156},
                {0.1250, 0.0156},
                {0.2083, 0.0156},
                {0.2500, 0.0156},
                {0.3333, 0.0156},
                {0.4167, 0.0156},
                {0.5000, 0.0156},
                {0.6667, 0.0156},
                {0.8333, 0.0156},
                {1.0000, 0.0156}
            };

            Dictionary<double, double> FP_Spacing = new Dictionary<double, double>()
            {
                {0.0417, 0.0156},
                {0.0625, 0.0156},
                {0.0833, 0.0156},
                {0.1042, 0.0156},
                {0.1250, 0.0156},
                {0.2083, 0.0156},
                {0.2500, 0.0156},
                {0.3333, 0.0156},
                {0.4167, 0.0156},
                {0.5000, 0.0156},
                {0.6667, 0.0156},
                {0.8333, 0.0156},
                {1.0000, 0.0156}
            };

            Dictionary<double, double> CuDWV_Spacing = new Dictionary<double, double>()
            {
                {0.0417, 0.0000},
                {0.0625, 0.0000},
                {0.0833, 0.0000},
                {0.1042, 0.0000},
                {0.1250, 0.0000},
                {0.2083, 0.0000},
                {0.2500, 0.0000},
                {0.3333, 0.0000},
                {0.4167, 0.0000},
                {0.5000, 0.0000},
                {0.6667, 0.0000},
                {0.8333, 0.0000},
                {1.0000, 0.0000}
            };

            AuthorList.Add(new PipeTypeInfo("KDS_Char_CI_NH", 10.00, CI_Spacing));
            AuthorList.Add(new PipeTypeInfo("KDS_CU_L", 20.00, CUL_Spacing));
            AuthorList.Add(new PipeTypeInfo("KDS_BLK", 21.00, BLK_GAL_Spacing));
            AuthorList.Add(new PipeTypeInfo("KDS_GALV", 21.00, BLK_GAL_Spacing));
            AuthorList.Add(new PipeTypeInfo("KDS_PVC", 20.00, PVC_spacing));
            AuthorList.Add(new PipeTypeInfo("KDS_FP", 20.00, FP_Spacing));
            AuthorList.Add(new PipeTypeInfo("KDS_CU_DWV", 20.00, CuDWV_Spacing));
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



 


    }  // end of Class InsertSleeve

    #region // PipeTypeInfo Class Settings
    class PipeTypeInfo
    {
        private string name;
        private double manufactureLength;
        private Dictionary<double, double> spacing_dictionary;

        public PipeTypeInfo(string name, double manufactureLength, Dictionary<double, double> d_sp_dct)
        {
            this.name = name;
            this.manufactureLength = manufactureLength;
            this.spacing_dictionary = d_sp_dct;
        }
        public string Name
        {
            get { return name; }
            set { name = value; }
        }
        public double ManufactureLength
        {
            get { return manufactureLength; }
            set { manufactureLength = value; }
        }
        public Dictionary<double, double> Spacing_Dictionary
        {
            get { return spacing_dictionary; }
            set { spacing_dictionary = value; }
        }
    }  // End Of Class PipeTypeInfo
    #endregion  // End of PipeTypeInfo




} // end KDS_Revit_Commands