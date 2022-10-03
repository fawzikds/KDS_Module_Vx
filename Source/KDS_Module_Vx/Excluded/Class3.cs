using System;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.IO;   // FOR StreamWriter
using System.Linq;
using System.Diagnostics;


namespace KDS_Module
{
    /// <summary>
    /// Description of GetSegSizes_CN.
    /// </summary>

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("23CF5F71-5468-438D-97C7-554F4F782936")]
    public class GetSegSizes_CN
    {
    	
    	const string _filename = "C:/Users/cad16/Desktop/pipesizes.txt";
 
  /*string FootToMmString( double a )
  {
    return Util.FootToMm( a )
      .ToString( "0.##" )
      .PadLeft( 8 );
  }
  */
 
  /// <summary>
  /// List all the pipe segment sizes in the given document.
  /// </summary>
  /// <param name="doc"></param>
  public void GetPipeSegmentSizes(Document doc )
  {
    FilteredElementCollector segments
      = new FilteredElementCollector( doc )
        .OfClass( typeof( Segment ) );
 
    using( StreamWriter file = new StreamWriter(
      _filename, true ) )
    {
      foreach( Segment segment in segments )
      {
        file.WriteLine( segment.Name );
 
        foreach( MEPSize size in segment.GetSizes() )
        {
          file.WriteLine( string.Format( "  {0} {1} {2}",
            size.NominalDiameter,
            size.InnerDiameter ,
            size.OuterDiameter   ));
            
        	/*file.WriteLine( string.Format( "  {0} {1} {2}",
            FootToMmString( size.NominalDiameter ),
            FootToMmString( size.InnerDiameter ),
            FootToMmString( size.OuterDiameter ) ) );
        	*/
        }
      }
    }
  }
  
  
  

        



}
}
