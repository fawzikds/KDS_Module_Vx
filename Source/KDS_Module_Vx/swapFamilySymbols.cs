/*
 * Created by SharpDevelop.
 * User: cad16
 * Date: 1/4/2021
 * Time: 11:26 AM
 */
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;




namespace KDS_Module
{
    /// <summary>
    /// Description of swapFamilies_CN find a family by a string match and replace with an equivalent base of angle criteria.
    /// 45 degrees  use a wye.
    /// 90degrees use a Combo..
    /// </summary>

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class swapFamilySymbols_CN : IExternalEventHandler
    {

        public static string GetFamilyName(Element e)
        {
            var eId = e?.GetTypeId();
            if (eId == null)
                return "";
            var elementType = e.Document.GetElement(eId) as ElementType;
            return elementType?.FamilyName ?? "";
        }


        public void Execute(UIApplication a)
        {

            UIDocument uidoc = a.ActiveUIDocument;
            Document actvDoc = uidoc.Document;
            /*
             * FilteredElementCollector dcs = new FilteredElementCollector(actvDoc)
                                                   .OfClass(typeof(FamilyInstance))
                                                   .OfCategory(BuiltInCategory
                                                   .OST_DetailComponents);
            */



            string selectString = "Cap - Welded - Generic";
            selectString = "Tee - Welded - Generic";
            const string param_name = "Angle 1";
            string replaceString = "Wye - Welded - Generic";
            //ElementId rplcE=null;
            ElementId rplcElmId = null;

            FamilySymbol rplcFamSymbol = null;


            var collector = new FilteredElementCollector(actvDoc);

            // Create a Filter to that finds elements of type "Family instance"
            ElementClassFilter famInstFilter = new ElementClassFilter(typeof(FamilyInstance));

            // Apply the filter to the elements in the active document 

            var allFamInst = from famInstElm in collector
                             .WherePasses(famInstFilter)
                                 //where element.Name.Contains(selectString)
                             select famInstElm;

            var allFamInstCount = allFamInst.ToList().Count;
            Autodesk.Revit.UI.TaskDialog.Show("swapFamilies_CN", "allFamInstCount = :  " + allFamInstCount);

            collector = new FilteredElementCollector(actvDoc);
            // Create a Filter to that finds elements of type "Family"
            ElementClassFilter famFilter = new ElementClassFilter(typeof(Family));
            var allFam = from rplcFamElm in collector
                             .WherePasses(famFilter)
                         select rplcFamElm;

            Autodesk.Revit.UI.TaskDialog.Show("swapFamilies_CN", "ALl families Count = :  " + allFam.ToList().Count);
            Family rplcFam = null;
            foreach (Family rf in allFam)
            {
                if (rf.Name.Contains(replaceString))
                {
                    rplcFam = rf; 
                    Autodesk.Revit.UI.TaskDialog.Show("swapFamilies_CN", "rf  Family Name = :  " + rplcFam.Name);
                    
                }
            }

            rplcFamSymbol = actvDoc.GetElement(rplcFam.GetFamilySymbolIds().FirstOrDefault()) as FamilySymbol;
            rplcFamSymbol.Activate();
            Autodesk.Revit.UI.TaskDialog.Show("swapFamilies_CN", "rplcFamSymbol Family Name = :  " + rplcFamSymbol.FamilyName);

            foreach (Autodesk.Revit.DB.FamilyInstance famInst in allFamInst)
            {
                try
                {
                    using (Transaction swapFamSym_trx = new Transaction(actvDoc, "swap FamilySymbol"))
                    {
                        swapFamSym_trx.Start();
                        //var famInst_el = actvDoc.GetElement(famInst.GetTypeId());
                        //Autodesk.Revit.UI.TaskDialog.Show("swapFamilies_CN", "famInst.Family = :  " + GetFamilyName(famInst_el));
                        if (selectString == famInst.Symbol.FamilyName)
                        {
                            Autodesk.Revit.UI.TaskDialog.Show("swapFamilies_CN", "famInst.Symbol.FamilyName = :  " + famInst.Symbol.FamilyName);

                            foreach (Parameter Param in famInst.Parameters)
                            {
                                if (param_name == Param.Definition.Name)
                                    //Autodesk.Revit.UI.TaskDialog.Show("swapFamilies_CN", "found one Angle 1 Parameter with value: " + Param.AsValueString());
                                    if (Param.AsValueString().Contains("45"))
                                    {
                                        Autodesk.Revit.UI.TaskDialog.Show("swapFamilies_CN", "found one Angle 1 Parameter with value: " + Param.AsValueString());
                                        Element e = actvDoc.GetElement(famInst.Id);
                                        e.ChangeTypeId(rplcFamSymbol.Id);
                                        //famInst.Symbol = rplcFamSymbol;
                                    }
                            }
                        }

                        swapFamSym_trx.Commit();
                    }//end transaction swapFamSym_trx

                }//end try
                catch (System.Exception e)
                {
                    Autodesk.Revit.UI.TaskDialog.Show("swapFamilies_CN", "I am in swapFamilies_CN.execte. Exception:  " + e.ToString());

                }// end catch
            }  // Foreach Family to edit

        }// end of Method FP_addFamilyparameters

        public string GetName()
        {
            return "External Event swapFamilySymbols";
        }



    }  // end of Class 

}  // namespace swapFamilies_NS







