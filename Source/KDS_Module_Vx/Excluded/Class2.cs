/*
 * Created by SharpDevelop.
 * User: Joshua.Lumley
 * Date: 2/10/2017
 * Time: 5:16 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Linq;
using System.Collections.Generic;
using System.Windows.Forms;



namespace KDS_Module
{
	/// <summary>
	/// Description of Class2.
	/// </summary>

	[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]	
   public class ButtonEE7Parameter : IExternalEventHandler  //this is the last when one making a checklist change, EE4 must be just for when an element is new
    {
        public void Execute(UIApplication uiapp)
        {
        	Document doc = uiapp.ActiveUIDocument.Document;

	   		try
            {
	   			
					//doc.ProjectInformation.GetParameters("Project Name")[0].Set("Space Elevator"); 			
	   			
	         		using (Transaction t = new Transaction(doc, "Modify Project Name"))
	                {
	                    t.Start();
	                       doc.ProjectInformation.GetParameters("Project Name")[0].Set("Space Elevator");
	                    t.Commit();
	                }
            }

            #region catch and finally
            catch (Exception ex)
            {
                TaskDialog.Show("Catch", "Failed due to:" + Environment.NewLine + ex.Message);
            }
            finally
            {
          
            }
            #endregion
        }


        public string GetName()
        {
            return "External Event Example";
        }
    }
	
	
}
