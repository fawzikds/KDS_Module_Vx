using System;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Diagnostics;


namespace KDS_Module
{
    /// <summary>
    /// Description of Class1.
    /// </summary>

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("23CF5F71-5468-438D-97C7-554F4F782936")]
    public class DrawLines
    {

        public void _99_DrawLinePatterns(bool patternsNotWeights, bool styles, UIDocument uidoc)
        {
            Document doc = uidoc.Document;
            bool modellines = false;
            bool closing = true;

            View view = doc.ActiveView;

            if (!modellines)
            {

                View3D myView3D = doc.ActiveView as View3D;

                if (myView3D != null)
                {

                    TaskDialog.Show("Not the correct type of view", "'Detail' lines can't be drawn in 3D view.");
                    return;
                }
            }

            if (modellines)
            {
                ViewPlan myViewPlan = doc.ActiveView as ViewPlan;

                if (myViewPlan == null)
                {

                    TaskDialog.Show("Not the correct type of view", "The active view must be a view 'plan'");
                    return;
                }
            }

            UIView uiview = null;

            IList<UIView> uiviews = uidoc.GetOpenUIViews();

            foreach (UIView uv in uiviews)
            {
                if (uv.ViewId.Equals(view.Id))
                {
                    uiview = uv;
                    break;
                }
            }


            View nextView = doc.GetElement(view.GetPrimaryViewId()) as View;

            if (nextView != null)
            {
                TaskDialog.Show("Cannot proceed", "Please goto the 'parent' view.");
                return;
            }


            Rectangle rect = uiview.GetWindowRectangle();
            IList<XYZ> corners = uiview.GetZoomCorners();
            XYZ pXchanges = corners[0];
            XYZ qXchanges = corners[1];


            double divideScreeenWidthBy50 = ((corners[1].X - corners[0].X) / 50);

            double halfOfX = corners[0].X;//corners[0].X + ((corners[1].X - corners[0].X)/2);
            pXchanges = new XYZ(halfOfX + (divideScreeenWidthBy50 * 10), corners[0].Y + (corners[0].Y / 10), 0);
            qXchanges = new XYZ(halfOfX + (divideScreeenWidthBy50 * 10), corners[1].Y - (corners[1].Y / 10), 0);

            XYZ[] pts = new XYZ[] {
                                    pXchanges,
                                    qXchanges,
                                    };

            List<Curve> profile = new List<Curve>(pts.Length); // 2013


            var collector = new FilteredElementCollector(doc).OfClass(typeof(LinePatternElement)).ToList();

            List<ElementId> ids = new List<ElementId>();
            for (int i = 0; i < collector.Count(); i++)
            {
                //	thenamesoflinepatterns = thenamesoflinepatterns + collector[i].Name + Environment.NewLine;

                ids.Add(collector[i].Id);
            }
            if (ids == null)
            {
                TaskDialog.Show("Cannot proceed", "is null");
                return;
            }

            // transaction.Commit();

            if (patternsNotWeights)
            {
                foreach (ElementId linePatternID in ids)
                {
                    notToRepeatLoopContents(closing, pts, profile, divideScreeenWidthBy50, doc, modellines, linePatternID, -1, patternsNotWeights, uidoc);
                }

            }
            else
            {
                for (int ii = 1; ii <= 16; ii++)
                {
                    notToRepeatLoopContents(closing, pts, profile, divideScreeenWidthBy50, doc, modellines, new ElementId(-1), ii, patternsNotWeights, uidoc);
                }
            }



            //writeDebug(thenamesoflinepatterns);
        }
        public void notToRepeatLoopContents(bool closing, XYZ[] pts, List<Curve> profile, double divideScreeenWidthBy50, Document doc, bool modellines, ElementId linePatternID, int ii, bool patternsNotWeights, UIDocument uidoc)
        {
            if (closing)
            {
                for (int i = 0; i <= pts.Length - 2; i++)
                {
                    profile.Add(Line.CreateBound(pts[i], pts[i + 1]));
                }
            }
            else
            {
                XYZ q = pts[pts.Length - 1];
                foreach (XYZ p in pts)
                {
                    profile.Add(Line.CreateBound(q, p));
                    q = p;
                }
            }

            for (int i = 0; i <= pts.Length - 1; i++)
            {
                pts[i] = new XYZ(pts[i].X + divideScreeenWidthBy50, pts[i].Y, pts[i].Z);
            }

            XYZ normal = XYZ.BasisZ;    // use basis of the z-axis (0,0,1) for normal vector
            XYZ origin = XYZ.Zero;  // origin is (0,0,0)  
            Plane geomPlane = Plane.CreateByNormalAndOrigin(normal, origin);

            foreach (Curve c in profile) // 2013
            {
                ModelCurve myModelCurve = null;
                DetailCurve myDetailCurve = null;


                if (modellines) myModelCurve = doc.Create.NewModelCurve(c, doc.ActiveView.SketchPlane);
                if (!modellines) myDetailCurve = doc.Create.NewDetailCurve(uidoc.ActiveView, c);

                /*	List<ElementId> lsArr = new List<ElementId>();

                   if (modellines) lsArr = myModelCurve.GetLineStyleIds().ToList();
                   if (!modellines) lsArr = myDetailCurve.GetLineStyleIds().ToList();*/

                OverrideGraphicSettings override2 = new OverrideGraphicSettings();// doc.ActiveView.GetElementOverrides(myDetailCurve.Id);
                if (patternsNotWeights)
                {
                    override2.SetProjectionLinePatternId(linePatternID);
                }
                else
                {
                    override2.SetProjectionLineWeight(ii);
                }
                doc.ActiveView.SetElementOverrides(myDetailCurve.Id, override2);

            }
            profile.Clear();
        }


        public void writeDebug(string x)
        {

            string path = (System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\` aa Subcct Details XLS Sheet");
            if (!System.IO.Directory.Exists(path)) System.IO.Directory.CreateDirectory(path);

            string subdirectory_reversedatetothesecond = (path + ("\\" + (DateTime.Now.ToString("yyyyMMddHHmmss"))));
            if (!System.IO.Directory.Exists(subdirectory_reversedatetothesecond)) System.IO.Directory.CreateDirectory(subdirectory_reversedatetothesecond);

            string FILE_NAME = (subdirectory_reversedatetothesecond + "\\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt");

            System.IO.File.Create(FILE_NAME).Dispose();

            System.IO.StreamWriter objWriter = new System.IO.StreamWriter(FILE_NAME, true);
            objWriter.WriteLine(x);
            objWriter.Close();
            Process.Start(FILE_NAME);
        }

        ICollection<ElementId> GetLineStyles(UIApplication m_rvtApp, Document m_rvtDoc)
        {
            Transaction tr =
                new Transaction(m_rvtDoc, "Get Line Styles");
            tr.Start();

            // Create a detail line

            View view = m_rvtDoc.ActiveView;
            XYZ pt1 = XYZ.Zero;
            XYZ pt2 = new XYZ(10.0, 0.0, 0.0);


            XYZ[] pts = new XYZ[] {
                    pt1,
                    pt2,
                    };

            List<Curve> profile = new List<Curve>(pts.Length); // 2013
            for (int ii = 0; ii <= pts.Length - 2; ii++)
            {
                profile.Add(Line.CreateBound(pts[ii], pts[ii + 1]));
            }

            DetailCurve dc = m_rvtDoc.Create.NewDetailCurve(view, profile[0]);
            ICollection<ElementId> lineStyles = dc.GetLineStyleIds();
            tr.RollBack();

            return lineStyles;
        }
    }



public class exportSegmentSizes_CN
{
	
}
}
