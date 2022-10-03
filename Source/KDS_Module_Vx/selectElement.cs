using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace KDS_Module
{
    public class SelectElement : IExternalCommand
    {
        /* User will need to adjust code to adapt to TYPE of object they wish to know info about, as well as those parameter (info) they wish to know
         */

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
			UIDocument uidoc = commandData.Application.ActiveUIDocument;

			Reference reference = uidoc.Selection.PickObject(ObjectType.Element);
			Pipe pipe = uidoc.Document.GetElement(reference) as Pipe;

            //XYZ p1 = inst?.FacingOrientation ?? new XYZ(0, 0, 0);
            //XYZ p2 = inst?.HandOrientation ?? new XYZ(0, 0, 0);
            string family = pipe?.Name ?? "NA";
            string id = pipe?.Id?.ToString() ?? "NA";

            LocationCurve lc = pipe.Location as LocationCurve;
            Curve c = lc.Curve;
            XYZ endpoint1 = c.GetEndPoint(0);
            XYZ endpoint2 = c.GetEndPoint(1);
            //string slope = inst.get_Parameter(BuiltInParameter.RBS_PIPE_SLOPE).AsValueString();

            TaskDialog.Show("test", "endpoint[0] : " + endpoint1 + "\nendpoint[1] : " + endpoint2 + "\npipe length: " + pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsValueString());
            //TaskDialog.Show("test", "family name: " + family);
            //TaskDialog.Show("test", "hand orientation: \n" + p2 + "\nfacing orientation: \n" + p1 + "\n\nfamily name: " + family + "\nid: " + id + "\nelement angle value: " + inst?.LookupParameter("Angle")?.AsDouble() * (180 / Math.PI) + "°" ?? "NA");
            return Result.Succeeded;

        }
    }
}
