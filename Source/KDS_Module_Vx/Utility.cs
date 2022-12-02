//
// (C) Copyright 2003-2019 by Autodesk, Inc.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE. AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//

using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.IO;
using System;
using System.Linq;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;

//namespace Revit.SDK.Samples.RoutingPreferenceTools.CS
namespace Utility
{


    #region  //est_data_class
    public class est_data_class
    {
        public object this[string propertyName] => GetType().GetProperty(propertyName)?.GetValue(this, null);

        public string Size { get; set; }
        public string famName { get; set; }
        public string ND1 { get; set; }
        public string ND2 { get; set; }
        public string ND3 { get; set; }
        public string ND4 { get; set; }
        public string KDS_MfrPart { get; set; }
        public string KDS_HPH { get; set; }
        public string KDS_MfrList { get; set; }
        public string KDS_MCAA_LBR_RATE { get; set; }
        public string KDS_LBR_RATE { get; set; }

        public static est_data_class FromCsv(string csvLine)
        {
            string[] values = csvLine.Split(',');
            est_data_class est_data = new est_data_class();

            est_data.Size = values[0];
            est_data.famName = values[1];
            est_data.ND1 = values[2];
            est_data.ND2 = values[3];
            est_data.ND3 = values[4];
            est_data.ND4 = values[5];
            est_data.KDS_MfrPart = values[6];
            est_data.KDS_HPH = values[7];
            est_data.KDS_MfrList = values[8];
            est_data.KDS_MCAA_LBR_RATE = values[9];
            est_data.KDS_LBR_RATE = values[10];

            return est_data;
        }  // End From Csv
        public override string ToString()
        {
            return "" +
                " Size: " + Size +
                " :: famName = " + famName +
                " :: ND1 = " + ND1 +
                " :: ND2 = " + ND2 +
                " :: ND3 = " + ND3 +
                " :: ND4 = " + ND4 +
                " :: KDS_MfrPart = " + KDS_MfrPart +
                " :: KDS_HPH = " + KDS_HPH +
                " :: KDS_MfrList = " + KDS_MfrList +
                " :: KDS_MCAA_LBR_RATE = " + KDS_MCAA_LBR_RATE +
                " :: KDS_LBR_RATE = " + KDS_LBR_RATE;
        }  // End of ToString

        public string ToFormatString()
        {
            return "" +
                " Size: " + Size +
                "\n famName = " + famName +
                "\n ND1 = " + ND1 +
                "\n ND2 = " + ND2 +
                "\n ND3 = " + ND3 +
                "\n ND4 = " + ND4 +
                "\n KDS_MfrPart = " + KDS_MfrPart +
                "\n KDS_HPH = " + KDS_HPH +
                "\n KDS_MfrList = " + KDS_MfrList +
                "\n KDS_MCAA_LBR_RATE = " + KDS_MCAA_LBR_RATE +
                "\n KDS_LBR_RATE = " + KDS_LBR_RATE;
        }  // End of ToFormatString
        public string ToCsvString()
        {
            return Size + "," +
                ND1 + "," + ND2 + "," + ND3 + "," + ND4 + "," +
                KDS_MfrPart + "," + KDS_HPH + "," + KDS_MfrList + "," + KDS_MCAA_LBR_RATE + "," + KDS_LBR_RATE;
        }  // End of ToCsvString

        public string ToCsvString_excld(string prop)
        {

            string props = this.ToCsvString();


            //return props.Replace(prop,"");
            return "," +
                ND1 + "," + ND2 + "," + ND3 + "," + ND4 + "," +
                KDS_MfrPart + "," + KDS_HPH + "," + KDS_MfrList + "," + KDS_MCAA_LBR_RATE + "," + KDS_LBR_RATE;
        }  // End of ToCsvString

        public string ToCsvString_wfname()
        {
            return Size + "," + famName + "," +
                ND1 + "," + ND2 + "," + ND3 + "," + ND4 + "," +
                KDS_MfrPart + "," + KDS_HPH + "," + KDS_MfrList + "," + KDS_MCAA_LBR_RATE + "," + KDS_LBR_RATE;
        }  // End of ToCsvString

    }  // End of Class est_data_class

    #endregion  //est_data_class


    #region  //famParam_Data_class
    // Data Class to hold the different pieces of a parameter, Name, type, default Value and Formula
    public class famParam_Data_class
    {
        public string Name { get; set; }
        public string type { get; set; }
        public string units { get; set; }
        public string dfltVlu { get; set; }
        public string formula { get; set; }

        public static famParam_Data_class FromCsv(string csvLine)
        {
            string[] values = csvLine.Split(',');
            famParam_Data_class famParam_Data = new famParam_Data_class();

            famParam_Data.Name = values[0];
            famParam_Data.type = values[1];
            famParam_Data.units = values[2];
            famParam_Data.dfltVlu = values[3];
            famParam_Data.formula = values[4];

            return famParam_Data;
        }  // End From Csv

        public override string ToString()
        {
            return
                " Name: " + Name +
                " :: type = " + type +
                " :: dfltVlu = " + dfltVlu +
                " :: units = " + units +
                " :: formula = " + formula;
        }  // End of ToString

        public string ToFormatString()
        {
            return
                " \n Name: " + Name +
                "\n type = " + type +
                "\n units = " + units +
                "\n dfltVlu = " + dfltVlu +
                "\n formula = " + formula;
        }  // End of ToFormatString

        public string ToCsvString()
        {
            return Name + "," + type + "," + units + "," + dfltVlu + "," + formula;
        }  // End of ToCsvString

    }  // End of Class est_data_class

    #endregion  //paramFormula_Data_class

    internal class Validation
    {
        public static bool ValidateMep(Autodesk.Revit.ApplicationServices.Application application)
        {
            return application.IsPipingEnabled;
        }
        public static void MepWarning()
        {
            Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceTools", "Revit MEP is required to run this addin.");
        }

        public static bool ValidatePipesDefined(Autodesk.Revit.DB.Document document)
        {
            FilteredElementCollector collector = new FilteredElementCollector(document);
            collector.OfClass(typeof(Autodesk.Revit.DB.Plumbing.PipeType));
            if (collector.Count() == 0)
                return false;
            else
                return true;
        }

        public static void PipesDefinedWarning()
        {
            Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceTools", "At least two PipeTypes are required to run this command.  Please define another PipeType.");
        }
    }

    internal class KDS_Convert
    {
        public static double ConvertValueDocumentUnits(double decimalFeet, Autodesk.Revit.DB.Document document)
        {

#if(CONFIG_R2019 || CONFIG_R2020)
            FormatOptions formatOption = document.GetUnits().GetFormatOptions(UnitType.UT_PipeSize);
         return UnitUtils.ConvertFromInternalUnits(decimalFeet, formatOption.DisplayUnits);
#else
            FormatOptions formatOption = document.GetUnits().GetFormatOptions(SpecTypeId.PipeSize);
            return UnitUtils.ConvertFromInternalUnits(decimalFeet, formatOption.GetUnitTypeId());
#endif
        }


        public static double ConvertValueToFeet(double unitValue, Autodesk.Revit.DB.Document document)
        {
            double tempVal = ConvertValueDocumentUnits(unitValue, document);
            double ratio = unitValue / tempVal;
            return unitValue * ratio;
        }
    }// KDS_Convert


    internal class KDS_Functions
    {

        private static List<string> FindAllRefrences(ref int ctr, string dir, string projectToSearch)
        {
            List<string> refs = new List<string>();
            foreach (var projFile in Directory.GetFiles(dir, "*.csproj", SearchOption.AllDirectories))
            {
                if (projFile.IndexOf(projectToSearch, StringComparison.OrdinalIgnoreCase) >= 0)
                    continue;

                var lines = File.ReadAllLines(projFile);

                foreach (var line in lines)
                {
                    if (line.IndexOf(projectToSearch, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        ctr++;
                        refs.Add(projFile);
                        break;
                    }
                }
            }

            return refs;
        }// End Of FindAllReferences


        #region // GetPipeLevel Function to get Level Name of Current Pipe //
        public string GetPipeLevel(UIDocument uidoc, Pipe p, string startEndMiddle)
        {
            string lvlName = null;
            LocationCurve lc = p.Location as LocationCurve;
            Curve c = lc.Curve;
            double zLocStart = c.GetEndPoint(0).Z;
            double zLocEnd = c.GetEndPoint(1).Z;
            double zLocMiddle = (zLocStart + zLocEnd) / 2;

            Autodesk.Revit.DB.Document actvDoc = uidoc.Document;


            List<Autodesk.Revit.DB.Level> sortedLevels_lst = new List<Autodesk.Revit.DB.Level>();

            List<Level> levels_lst = new FilteredElementCollector(actvDoc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements().Select(el => el as Level).ToList();

            if (levels_lst.Count > 0)
            {
                //sortedLevels_lst = levels_lst.OrderByDescending(ordr => ordr.Elevation).ToList();
                sortedLevels_lst = levels_lst.OrderBy(ordr => ordr.Elevation).ToList();
            }
            /* For Debug Only.--- List Levels and Lowest Floor
                       if (levels_lst.Count > 0)
                       {
                               string tsd_lvl = "";
                               tsd_lvl = "levels Found:\n";
                               int int_lvl = 0;
                               foreach (Autodesk.Revit.DB.Level srLvl in sortedLevels_lst)
                               {
                                   int_lvl++;
                                   tsd_lvl += int_lvl + "- " + srLvl.Name + " @ Z: " + srLvl.Elevation +"\n";
                                   //levels_lst.Add(level);

                               }
                               tsd_lvl += " Lowest Level is: " + sortedLevels_lst[0].Name;
                               TaskDialog.Show("sdsfgsd", tsd_lvl);
                        }
            */


            #region // User wants Z-Coordinate of Start EndPoint for Pipe //

            if (startEndMiddle == "start" || startEndMiddle == "Start")
            {
                foreach (Autodesk.Revit.DB.Level level in sortedLevels_lst)
                {
                    if (zLocStart > level.Elevation)
                    {
                        lvlName = level.Name;
                    }
                }

                if (lvlName == null)
                {
                    lvlName = sortedLevels_lst.First().Name;
                }
            }
            #endregion

            #region // User wants Z-Coordinate of End Endpoint for Pipe //
            if (startEndMiddle == "end" || startEndMiddle == "End")
            {
                foreach (Autodesk.Revit.DB.Level level in sortedLevels_lst)
                {
                    if (zLocEnd > level.Elevation)
                    {
                        lvlName = level.Name;
                    }
                }

                if (lvlName == null)
                {
                    lvlName = sortedLevels_lst.First().Name;
                }
            }
            #endregion

            #region // User wants Z-Coordinate of Midpoint of Pipe //
            if (startEndMiddle == "middle" || startEndMiddle == "Middle")
            {
                foreach (Autodesk.Revit.DB.Level level in sortedLevels_lst)
                {
                    if (zLocMiddle > level.Elevation)
                    {
                        lvlName = level.Name;
                    }
                }

                if (lvlName == null)
                {
                    lvlName = sortedLevels_lst.First().Name;
                }
            }
            #endregion

            else { TaskDialog.Show("error", "value for startEndMiddle is not \"start\", \"end\", or \"middle\""); }

            return lvlName;
        }
        #endregion

    }

}
