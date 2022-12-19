using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KDS_Module_Vx
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class SwapFamilies : IExternalCommand
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
                MainContent = "Select type of fitting to change",
                CommonButtons = TaskDialogCommonButtons.Close,
                DefaultButton = TaskDialogResult.Close,
                FooterText = "KDS Plumbing and Heating Services",
                TitleAutoPrefix = false,
                MainIcon = TaskDialogIcon.TaskDialogIconInformation,
            };
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Replace Wye With Santee and Combo");
            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Replace shortsweep or QuarterBend");

            TaskDialogResult tResult = mainDialog.Show();

            if (TaskDialogResult.CommandLink1 == tResult)
            {
                ChangeWye(uidoc.Document);
            }
            if (TaskDialogResult.CommandLink2 == tResult)
            {
                ChangeElbow(uidoc.Document);
            }
            else
            { return; }
        } // end DialogBox
        #endregion

        #region // ChangeWye Function to change Wyes to Santee or Combo //
        public void ChangeWye(Document actvDoc)
        {
            #region  // Define Family Names for collector search.

            string selectString = "KDS_Char_CI_NH-Wye"; // NAME OF OBJECT TO BE REPLACED //
            string sanTeeString = "KDS_Char_CI_NH-SanTee"; // NAME OF REPLACING OBJECT 
            string comboString = "KDS_Char_CI_NH-Combo"; // NAME OF REPLACING OBJECT 
            const string param_name = "Angle"; // NAME OF PARAMETER BEING COMPARED / ANALYZED //
            double angle = Math.PI / 2;
            double angleTolerance = 0.15;
            double lowerAngleLimit = angle - angle * angleTolerance;
            double upperAngleLimit = angle + angle * angleTolerance;

            //TaskDialog.Show("test", "angle: " + angle * (180 / Math.PI) + "°\nlowerAngleLimit: " + lowerAngleLimit * (180 / Math.PI) + "°\nupperAngleLimit: " + upperAngleLimit * (180 / Math.PI) + "°");

            var watch = System.Diagnostics.Stopwatch.StartNew();

            #endregion  // Define Family Names for collector search.

            #region // Get current family to be replaced based on its name

            FilteredElementCollector currCollector = new FilteredElementCollector(actvDoc);
            Family currFam = null;
            currFam = currCollector.OfClass(typeof(Family)).OfType<Family>().FirstOrDefault(f => f.Name.Equals(selectString));

            FamilySymbol currFamSymbol = null;
            currFamSymbol = actvDoc.GetElement(currFam.GetFamilySymbolIds().FirstOrDefault()) as FamilySymbol;
            currFamSymbol.Activate();
            #endregion

            #region // Get replacement Families and Symbols based on its name.

            FilteredElementCollector rplcCollector = new FilteredElementCollector(actvDoc);

            Family sanTeeFam = rplcCollector.OfClass(typeof(Family)).OfType<Family>().FirstOrDefault(f => f.Name.Equals(sanTeeString));
            Family comboFam = rplcCollector.OfClass(typeof(Family)).OfType<Family>().FirstOrDefault(f => f.Name.Equals(comboString));

            FamilySymbol sanTeeFamSymbol = actvDoc.GetElement(sanTeeFam.GetFamilySymbolIds().FirstOrDefault()) as FamilySymbol;
            sanTeeFamSymbol.Activate();

            FamilySymbol comboFamSymbol = actvDoc.GetElement(comboFam.GetFamilySymbolIds().FirstOrDefault()) as FamilySymbol;
            comboFamSymbol.Activate();

            #endregion //Get the replacement Family  based on its name. 

            #region // Get All Family instances to be replaced in Project

            FilteredElementCollector instCollector = new FilteredElementCollector(actvDoc);

            List<FamilyInstance> allFamInst = instCollector.OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                                                                                            .Where(inst => inst.Symbol.Id.Equals(currFamSymbol.Id) && 
                                                                                                   inst.LookupParameter(param_name).AsDouble() >= lowerAngleLimit &&
                                                                                                   inst.LookupParameter(param_name).AsDouble() <= upperAngleLimit)
                                                                                            .ToList();
            #endregion

            #region // Filtering based on orientation //

            int s = 0;
            int c = 0;
            double filterTolerance = 0.259;

            foreach (FamilyInstance inst in allFamInst)
            {
                using (Transaction swapFamSym_trx = new Transaction(actvDoc, "swap FamilySymbol"))
                {
                    swapFamSym_trx.Start();

                    if (inst.HandOrientation.Z > (1 - filterTolerance) && inst.FacingOrientation.Z < (1 + filterTolerance))
                    {
                        if (inst.FacingOrientation.Z > (0 - filterTolerance) && inst.FacingOrientation.Z < (0 + filterTolerance))
                        {
                            Element e = actvDoc.GetElement(inst.Id);
                            e.ChangeTypeId(sanTeeFamSymbol.Id);
                            s++;
                        }
                        else
                        {
                            Element e = actvDoc.GetElement(inst.Id);
                            e.ChangeTypeId(comboFamSymbol.Id);
                            c++;
                        }
                    }
                    else
                    {
                        Element e = actvDoc.GetElement(inst.Id);
                        e.ChangeTypeId(comboFamSymbol.Id);
                        c++;
                    }
                    swapFamSym_trx.Commit();
                }
            }
            watch.Stop();
            var elapsedTime = watch.ElapsedMilliseconds;
            TaskDialog.Show("test", "# of fittings changed to Santees: " + s + "\n# of fittings changed to combos: " + c + "\n execution time : " + elapsedTime / 1000.0 + " seconds");

            #endregion

        }// ChangeFitting Function End
        #endregion

        #region // ChangeElbow Function //
        public void ChangeElbow(Document actvDoc)
        {
            #region  // Define Constants
            double angle = Math.PI / 2;   // To compare to ehe Angle of a Bend
            double angleTolerance = 0.15;   // Tolerance to comparison of angle of a bend instance with the one to replace with.
            double lowerAngleLimit = angle - angle * angleTolerance;   // Calculated angles with Tolerance
            double upperAngleLimit = angle + angle * angleTolerance;   // Calculated angles minus Tolerance.
            double isPipeVertical_Tolerance = 0.25;   // Needed to decide whether a pipe is condiered vertical or not. This is for choosing Bends on Vertical pipe

            // List of Families to Be Swapped.  in this case they are all the same.
            List<string> toBeSwappedFamName_lst = new List<string>() {
                "KDS_Char_CI_NH-Bend_90_72_60_45_22.5",
                "KDS_Char_CI_NH-Bend_90_72_60_45_22.5",
                "KDS_Char_CI_NH-Bend_90_72_60_45_22.5",
                "KDS_Char_CI_NH-Bend_90_72_60_45_22.5",
                "KDS_Char_CI_NH-Bend_90_72_60_45_22.5"
            };
            // List of parameter to be used in sorting out the Bends, 90, 72, 60, 45, 22,5.
            // In our Bend families, we are using an "Angle 1" Parameter which is a stepped value of the "Angle" parameter.
            // So essentially we do not need to use this tolerance to angles.  but kept it in case other families did not have that.
            List<string> swappingFamCond0_lst = new List<string>() {
                "Angle 1",
                "Angle 1",
                "Angle 1",
                "Angle 1",
                "Angle 1",
            };

            List<double> swappingFamCond1_lst = new List<double>() {
                Math.PI * 2/16,   // 22.5
                Math.PI * 2/8,    // 45
                Math.PI * 2/6,    // 60
                Math.PI * 2/5,    // 72
                Math.PI * 2/4     // 90
            };

            List<double> swappingFamCond2_lst = new List<double>() {
                Math.PI * 2/16 * (1 - angleTolerance),
                Math.PI * 2/8  * (1 - angleTolerance),
                Math.PI * 2/6  * (1 - angleTolerance),
                Math.PI * 2/5  * (1 - angleTolerance),
                Math.PI * 2/4  * (1 - angleTolerance)
            };

            List<double> swappingFamCond3_lst = new List<double>() {
                Math.PI * 2/16 * (1 + angleTolerance),
                Math.PI * 2/8  * (1 + angleTolerance),
                Math.PI * 2/6  * (1 + angleTolerance),
                Math.PI * 2/5  * (1 + angleTolerance),
                Math.PI * 2/4  * (1 + angleTolerance)
            };

            List<bool> swappingFamAlternate1_lst = new List<bool>() {
                false,
                false,
                false,
                false,
                true,
            };

            // These are 2 lists of families to swap with.. Based on the alternate value above, if it is true, then we will use the hardcoded condition to replace with another family when met.
            List<string> swapingFamName1_lst = new List<string>() {
                "KDS_Char_CI_NH-SixteenthBend",
                "KDS_Char_CI_NH_EighthBend",
                "KDS_Char_CI_NH-SixthBend",
                "KDS_Char_CI_NH-FifthBend",
                "KDS_Char_CI_NH-ShortSweep",
            };

            List<string> swapingFamName2_lst = new List<string>() {
                "KDS_Char_CI_NH-SixteenthBend",
                "KDS_Char_CI_NH-EighthBend",
                "KDS_Char_CI_NH-SixthBend",
                "KDS_Char_CI_NH-FifthBend",
                "KDS_Char_CI_NH-QuarterBend",
                };

            List<int> swapingFamCnt0_lst = new List<int>()
            {
                0,
                0,
                0,
                0,
                0
            };

            List<int> swapingFamCnt1_lst = new List<int>()
            {
                0,
                0,
                0,
                0,
                0
            };


            #endregion // Define Constants

            #region   // Create a Rules List to swap families
            // create a list from the lists above using zip.
            // the final structure is (Tbs, c0, c1, c2,c3, a1, sf1,sf2)  accessed using dot notation swaprule.tbs, etc
            var swapRules_lst_Zip0 = toBeSwappedFamName_lst.Zip(swappingFamCond0_lst.Zip(swappingFamCond1_lst.Zip(swappingFamCond2_lst.Zip(swappingFamCond3_lst.Zip(swappingFamAlternate1_lst.Zip(swapingFamName1_lst.Zip(swapingFamName2_lst,
                (sf1, sf2) => new { sf1, sf2 }),
                (a1, sf2sf1) => new { a1, sf2sf1.sf1, sf2sf1.sf2 }),
                (c3, a1sf2sf1) => new { c3, a1sf2sf1.a1, a1sf2sf1.sf1, a1sf2sf1.sf2 }),
                (c2, a1sf2sf1c3) => new { c2, a1sf2sf1c3.a1, a1sf2sf1c3.sf1, a1sf2sf1c3.sf2, a1sf2sf1c3.c3 }),
                (c1, a1sf2sf1c3c2) => new { c1, a1sf2sf1c3c2.a1, a1sf2sf1c3c2.sf1, a1sf2sf1c3c2.sf2, a1sf2sf1c3c2.c2, a1sf2sf1c3c2.c3 }),
                (c0, a1sf2sf1c3c2c1) => new { c0, a1sf2sf1c3c2c1.a1, a1sf2sf1c3c2c1.sf1, a1sf2sf1c3c2c1.sf2, a1sf2sf1c3c2c1.c1, a1sf2sf1c3c2c1.c2, a1sf2sf1c3c2c1.c3 }),
                (tbs, a1sf2sf1c3c2c1c0) => new { tbs, a1sf2sf1c3c2c1c0.a1, a1sf2sf1c3c2c1c0.sf1, a1sf2sf1c3c2c1c0.sf2, a1sf2sf1c3c2c1c0.c0, a1sf2sf1c3c2c1c0.c1, a1sf2sf1c3c2c1c0.c2, a1sf2sf1c3c2c1c0.c3 });


            var   swapRules_lst_Zip = toBeSwappedFamName_lst.Zip(
                swappingFamCond0_lst.Zip(swappingFamCond1_lst.Zip(swappingFamCond2_lst.Zip(swappingFamCond3_lst.Zip(
                    swappingFamAlternate1_lst.Zip(swapingFamName1_lst.Zip(swapingFamName2_lst.Zip(
                        swapingFamCnt0_lst.Zip(swapingFamCnt1_lst,
    (cnt0, cnt1) => new { cnt0, cnt1 }),
    (sf2, g0) => new { sf2, g0.cnt0, g0.cnt1 }),
    (sf1, g1) => new { sf1, g1.sf2, g1.cnt0, g1.cnt1 }),
    (a1, g2) => new { a1, g2.sf1, g2.sf2, g2.cnt0, g2.cnt1 }),
    (c3, g3) => new { c3, g3.a1, g3.sf1, g3.sf2, g3.cnt0, g3.cnt1 }),
    (c2, g4) => new { c2, g4.c3, g4.a1, g4.sf1, g4.sf2, g4.cnt0, g4.cnt1 }),
    (c1, g5) => new { c1, g5.c2, g5.c3, g5.a1, g5.sf1, g5.sf2, g5.cnt0, g5.cnt1 }),
    (c0, g6) => new { c0, g6.c1, g6.c2, g6.c3, g6.a1, g6.sf1, g6.sf2, g6.cnt0, g6.cnt1 }),
    (tbs, g7) => new { tbs, g7.c0, g7.c1, g7.c2, g7.c3, g7.a1, g7.sf1, g7.sf2, g7.cnt0, g7.cnt1 }
    );


            //List<SwapRules> swapRules_lst2 = new List<SwapRules>();


            #endregion  //  Create a Rules List to swap families

            #region // Print out Swap Rules  --- Debug Purposes
            string swapStr = null;

            foreach (var swapr in swapRules_lst_Zip)
            {
                swapStr +=
                    "\nswapr.tbs: " + swapr.tbs +
                    "\nswapr.c0: " + swapr.c0 +
                     "\nswapr.c1: " + swapr.c1 * (180 / Math.PI) +
                    "°\nswapr.c2: " + swapr.c2 * (180 / Math.PI) +
                    "°\nswapr.c3: " + swapr.c3 * (180 / Math.PI) +
                    "\nswapr.tbs: " + swapr.a1 +
                    "\nswapr.sf1: " + swapr.sf1 +
                    "\nswapr.sf1: " + swapr.sf2 +
                    "\nswapr.cnt0: " + swapr.cnt0 +
                    "\nswapr.cnt1: " + swapr.cnt1;
            }


            TaskDialog.Show("swapTest", swapStr);
            //string tbs, strng c0, double c1, double c2, double c3, bool a1, string sf1, string sf2, int cnt0, int cnt1

            //int t1 = swapRules_lst_Zip.FirstOrDefault();


            #endregion   // Print out Swap Rules


            #region // Filter based on orientation Then Make the swap
            var watch = System.Diagnostics.Stopwatch.StartNew();

            foreach (var swapr in swapRules_lst_Zip)
            {
                FamilySymbol tbsFamSymb = get_FamSymbol(actvDoc, swapr.tbs);
                tbsFamSymb.Activate();

                List<FamilyInstance> tbsFamInst_lst = get_FamInst_lst(actvDoc, tbsFamSymb, swapr.c0, swapr.c2, swapr.c3);
                if (tbsFamInst_lst.Count == 0) { TaskDialog.Show("swapr", "Count of " + swapr.tbs + " is: " + tbsFamInst_lst.Count + "! Skipping Loop."); continue; }

                FamilySymbol famSymb1 = get_FamSymbol(actvDoc, swapr.sf1);
                if (famSymb1 == null) { TaskDialog.Show("swapr", "There is no symbol for:  " + swapr.sf1 + " ! Skipping Loop."); continue; }

                int cnt0 = 0;
                int cnt1 = 0;
                
                foreach (FamilyInstance inst in tbsFamInst_lst)
                {
                    using (Transaction swapFamSymb_trx = new Transaction(actvDoc, "swap FamilySymbol"))
                    {
                        swapFamSymb_trx.Start();
                        try
                        {
                            if (swapr.a1)
                            {
                                FamilySymbol famSymb2 = get_FamSymbol(actvDoc, swapr.sf2);
                                if (famSymb2 == null) { TaskDialog.Show("swapr", "There is no symbol for:  " + swapr.sf2 + " ! Skipping Loop."); continue; }
                                if (vert2horz(actvDoc, inst, isPipeVertical_Tolerance))
                                {
                                    swapFams(actvDoc, famSymb1, inst);
                                    cnt0++;
                                    //swapr.cnt0.Equals(cnt0);
                                }
                                else
                                {
                                    swapFams(actvDoc, famSymb2, inst);
                                    cnt1++;
                                    //swapr.cnt1.Equals(cnt1);
                                }
                            }
                            else
                            {
                                swapFams(actvDoc, famSymb1, inst);
                                cnt0++;
                                swapr.cnt0.Equals(cnt0);
                            }

                            swapFamSymb_trx.Commit();

                        }
                        catch (System.Exception ex)
                        { TaskDialog.Show("swapRules", "Exception: " + ex); }
                    }  // End of using swapFamSymb_trx
                }  // End of all family instances loop

                watch.Stop();
                var elapsedTime = watch.ElapsedMilliseconds;
                TaskDialog.Show("swapFmilies", "\n # of fittings changed to " + swapr.sf1 + ": " + swapr.cnt0 +
                                               "\n # of fittings changed to " + swapr.sf2 + ": " + swapr.cnt1 +
                                               "\n in : " + elapsedTime / 1000.0 + " seconds");
            }  // End of all Rules loop
            #endregion  // Filter based on orientation Then Make the swap 

        } //Change Elbow Function End
        #endregion




        public FamilySymbol get_FamSymbol(Document actvDoc, string famName)
        {
            FilteredElementCollector actvDocCollector = new FilteredElementCollector(actvDoc);

            Family Fam = actvDocCollector.OfClass(typeof(Family)).OfType<Family>().FirstOrDefault(f => f.Name.Equals(famName));

            FamilySymbol FamSymbol = actvDoc.GetElement(Fam.GetFamilySymbolIds().FirstOrDefault()) as FamilySymbol;
            //shortSweepFamSymbol.Activate();
            return FamSymbol;
        }

        public List<FamilyInstance> get_FamInst_lst(Document actvDoc, FamilySymbol genericBendFamSymbol, string param_name, double lowerAngleLimit, double upperAngleLimit)
        {
            FilteredElementCollector instCollector = new FilteredElementCollector(actvDoc);

            List<FamilyInstance> allFamInst = instCollector.OfClass(typeof(FamilyInstance)).Cast<FamilyInstance>()
                                                                                            .Where(inst => inst.Symbol.Id.Equals(genericBendFamSymbol.Id) && inst.LookupParameter(param_name)
                                                                                            .AsDouble() >= lowerAngleLimit && inst.LookupParameter(param_name).AsDouble() <= upperAngleLimit).ToList();
            return allFamInst;
        }








        public bool vert2horz(Document actvDoc, FamilyInstance inst, double filterTolerance)
        {
            bool changeToShortSweep = true;

            if (inst.HandOrientation.Z > (0 - filterTolerance) && inst.HandOrientation.Z < (0 + filterTolerance))
            {
                if (inst.FacingOrientation.Z > (-1 - filterTolerance) && inst.FacingOrientation.Z < (-1 + filterTolerance))
                {
                    changeToShortSweep = false;
                }
            }
            else
            {
                changeToShortSweep = true;
            }

            if (inst.HandOrientation.Z > (1 - filterTolerance) && inst.HandOrientation.Z < (1 + filterTolerance))
            {
                if (inst.FacingOrientation.Z > (0 - filterTolerance) && inst.FacingOrientation.Z < (0 + filterTolerance))
                {
                    changeToShortSweep = false;
                }
            }
            else
            {
                changeToShortSweep = true;
            }

            return changeToShortSweep;
        }  // end of vert2horz







        public void swapFams(Document actvDoc, FamilySymbol famSymb, FamilyInstance famInst)
        {
            famSymb.Activate();
            Element e = actvDoc.GetElement(famInst.Id);
            e.ChangeTypeId(famSymb.Id);
        }  // End of swapFams



    }  // End of Class SwapFamilies IExternalCommand

    public class SwapRules
    {
        public string toBeSwappedName { get; set; }
        public string cond0 { get; set; }
        public double cond1 { get; set; }
        public double cond2 { get; set; }
        public double cond3 { get; set; }
        public bool alt1 { get; set; }
        public string sf1 { get; set; }
        public string sf2 { get; set; }
        public int cnt0 { get; set; }
        public int cnt1 { get; set; }
           
}





}  // End of namespace KDS_Module