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
    /// Description of Class1.
    /// </summary>

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("23CF5F71-5468-438D-97C7-554F4F782936")]
    
        public class GetPipeTypes_CN
    {
    	
    	const string _filename = "C:/Users/cad16/Desktop/pipeTypes.txt";
 
 
  /// <summary>
  /// List all the pipe segment sizes in the given document.
  /// </summary>
  /// <param name="doc"></param>
  public void GetPipeTypes(Document doc )
  {
  	var PipeTypes_colc = new FilteredElementCollector(doc);
  	
  	var PipeTypes_el = PipeTypes_colc.OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsElementType().ToElements();

  	
 
    using( StreamWriter file = new StreamWriter(
      _filename, true ) )
    {
      foreach( Element pt in PipeTypes_el )
      {
        file.WriteLine( pt.Name );
 
      }
    }
  }
  

}
        
        
    
    
    public class GetPipingSystem_CN
    {
    	
    	const string _filename = "C:/Users/cad16/Desktop/PipingSystems.txt";
 
 
  /// <summary>
  /// List all the pipe segment sizes in the given document.
  /// </summary>
  /// <param name="doc"></param>
  public void GetPipingSystem(Document doc )
  {
  	var PipingSystems_colc = new FilteredElementCollector(doc);
  	
  	var PipingSystems_el = PipingSystems_colc.OfCategory(BuiltInCategory.OST_PipingSystem).ToElements();
    

      using( StreamWriter file = new StreamWriter(
      _filename, true ) )
    {
      foreach( Element ps in PipingSystems_el )
      {
        file.WriteLine( ps.Name );
 
      }
    }
  }
  

}
}
