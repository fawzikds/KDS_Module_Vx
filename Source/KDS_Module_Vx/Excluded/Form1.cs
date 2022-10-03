/*
 * Created by SharpDevelop.
 * User: Joshua.Lumley
 * Date: 8/09/2017
 * Time: 11:45 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.IO;
using System.Windows.Forms;
//using System.IO;
using System.Xml;
//using Autodesk.Revit.DB;
using System.Xml.Linq;
using SAMPS = Revit.SDK.Samples.RoutingPreferenceTools.CS;



namespace KDS_Module
{
    /// <summary>
    /// Description of Form1.
    /// </summary>
    public partial class Form1 : System.Windows.Forms.Form
    {

 

        public Document doc { get; set; }

 

 


        //form_CommandReadPreferences_CN CommandReadPreferences_obj;
        CommandReadPreferences_CN CommandReadPreferences_obj;
        ExternalEvent CommandReadPreferences_ev;


        rplcFamilyInstances_CN rplcFamilyInstances_obj;
        ExternalEvent rplcFamilyInstances_ev;

        //swapFamilySymbols_CN swapFamilySymbols_obj;
        //ExternalEvent swapFamilySymbols_ev;


      
        public Form1()
        {
            //
            // The InitializeComponent() call is required for Windows Forms designer support.
            //
            InitializeComponent();

 

 

            CommandReadPreferences_obj = new CommandReadPreferences_CN();
            CommandReadPreferences_ev = ExternalEvent.Create(CommandReadPreferences_obj);

            rplcFamilyInstances_obj = new rplcFamilyInstances_CN();
            rplcFamilyInstances_ev = ExternalEvent.Create(rplcFamilyInstances_obj);

 
        }



 

        void Button3Click(object sender, EventArgs e)
        {
            // This Throw Command is just to illistrate the point that if we put our code here we may see the unhandled exception.
            // So Commands Should be within a transaction which is within a try command
            // throw new InvalidOperationException();

            try
            {
                // This Throw Command is just to illistrate the point that if we put our code here we may see the unhandled exception.
                // So Commands Should be within a transaction which is within a try command
                //throw new InvalidOperationException();  
                /*using (Transaction t = new Transaction(doc, "Set a parameters"))
		              {
		                  t.Start();
		                                              doc.ProjectInformation.GetParameters("Project Name")[0].Set("Space Elevator");  //this needs to change in two places
		                   t.Commit();
		               }
		      	*/


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

 



        void Button7Click(object sender, EventArgs e)
        {
            CommandReadPreferences_ev.Raise();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            //swapFamilySymbols_ev.Raise();
            rplcFamilyInstances_ev.Raise();
        }
    }




    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("23CF5F71-5468-438D-97C7-554F4F782936")]




 







    public class CommandReadPreferences_CN : IExternalEventHandler
    {

        public void Execute(UIApplication a)
        {

            UIDocument uidoc = a.ActiveUIDocument;
            Document doc = uidoc.Document;

            var fileContent = string.Empty;
            var filePath = string.Empty;

            OpenFileDialog openFileDialog = new OpenFileDialog();

        //openFileDialog.InitialDirectory = "c:\\";

        //openFileDialog.InitialDirectory = "C:\\Users\\cad16\\Desktop";
            openFileDialog.InitialDirectory = "C:\\ProgramData\\Autodesk\\Revit\\Macros\\2020\\Revit\\AppHookup\\KDS_Module\\Source\\KDS_Module";
            openFileDialog.DefaultExt = ".xml";
            openFileDialog.Filter = "RoutingPreference Builder Xml files (*.xml)|*.xml";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "Just Read the File.  " + openFileDialog.FileName);

                StreamReader reader = new StreamReader(openFileDialog.FileName);
                XDocument routingPreferenceBuilderDoc = XDocument.Load(new XmlTextReader(reader));
                reader.Close();
                Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "Just Loaded XmlTextReader  " + openFileDialog.FileName);

                //Distribute the .xsd file to routing preference builder xml authors as necessary.
                string xmlValidationMessage;
                if (!SAMPS.SchemaValidationHelper.ValidateRoutingPreferenceBuilderXml(routingPreferenceBuilderDoc, out xmlValidationMessage))
                {
                    Autodesk.Revit.UI.TaskDialog.Show("Form1", "Xml file is NOT a valid RoutingPreferenceBuilder xml document.  Please check RoutingPreferenceBuilderData.xsd.  " + xmlValidationMessage);
                    return;
                }
                try
                {
                    //Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "Routing Preferences Trying to Build.");
                    SAMPS.RoutingPreferenceBuilder builder = new SAMPS.RoutingPreferenceBuilder(doc);
                    //Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "Routing Preferences Just Created new Builder.");
                    builder.ParseAllPipingPoliciesFromXml(routingPreferenceBuilderDoc);
                    //Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "Routing Preferences Just came back from build.");
                }
                catch (SAMPS.RoutingPreferenceDataException ex)
                {
                    Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder error: ", ex.ToString());
                }
            }
            return;

        }

        public string GetName()
        {
            return "External Event CommandReadPreferences_CN";
        }


    }


}
