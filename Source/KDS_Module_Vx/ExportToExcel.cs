using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace KDS_Module
{
    public class ExportToExcel : IExternalCommand
    {
        #region // Execute Region of Code (Invoked by KDS Ribbon) //
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Lists for Headers in Excel File. Divided into categories based on type of parameter //
            List<string> elm_hdr_lst = new List<string> { "ElementID", "Family", "Type" };
            List<string> fam_hdr_lst = new List<string> { "System Classification", "KDS_MCAA_LBR_RATE", "KDS_LBR_RATE", "KDS_HPH", "KDS_MfrList", "KDS_MfrPart", "Category", "Size", "Length", "System Type" };
            List<string> calc_hdr_lst = new List<string> { "Is Vertical", "Level", "Diameter", "System Name" };

            //Calling of functions //
            LoadReference();
            DialogBox(commandData.Application.ActiveUIDocument, elm_hdr_lst, fam_hdr_lst, calc_hdr_lst);

            return Result.Succeeded;
        }
        #endregion

        #region // Load Reference Function //
        [STAThread]
        static void LoadReference()
        {
            string resource1 = "KDS_Revit_Commands.DocumentFormat.OpenXml.dll"; // ClassName.EmbeddedResource
            EmbeddedAssembly.Load(resource1, "DocumentFormat.OpenXml.dll"); // Raises Event

            AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve); // Event Handler
        }

        static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            return EmbeddedAssembly.Get(args.Name); // event, gets embedded resource
        }
        #endregion 

        #region // DialogBox Function to handle Facilitation of Function //
        public void DialogBox(UIDocument uidoc, List<string> elmHdrList, List<string> famHdrList, List<string> calcHdrList)
        {
            #region // DialogBox Settings //
            TaskDialog mainDialog = new TaskDialog("Export to Excel")
            {
                MainInstruction = "Please choose an option",
                MainContent = "Choose an option below",
                Title = "Export To Excel",
                CommonButtons = TaskDialogCommonButtons.Close,
                DefaultButton = TaskDialogResult.Close,
                FooterText = "KDS Plumbing and Heating Services",
                MainIcon = TaskDialogIcon.TaskDialogIconInformation,
                TitleAutoPrefix = false,
            };

            mainDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Export to Excel");

            TaskDialogResult tResult = mainDialog.Show(); // returns the users response from maindialog

            #endregion // DialogBox settings

            #region // Set up of Logic to respond to user's main dialog choice //

            bool Yes;

            if (TaskDialogResult.CommandLink1 == tResult)
            { Yes = true; }
            else
            { return; }
            #endregion

            #region // User Chose Export to Excel Option //
            if (Yes) // User Chose Export To Excel
            {
                // Collection of Plumbing Elements Lists //
                List<Pipe> pipes = GetPipes(uidoc);
                List<FamilyInstance> fittings = GetPipeFittings(uidoc);
                List<FamilyInstance> fixtures = GetPlumbingFixtures(uidoc);

                // creation of header titles for Excel File //
                List<string> hdrList = new List<string>();
                hdrList.AddRange(elmHdrList);
                hdrList.AddRange(famHdrList);
                hdrList.AddRange(calcHdrList);

                // Initialization for Excel Sheet //
                string filePath = "C:\\Users\\Jake\\Desktop\\Excel Files\\Elements.xltm";
                string sheetName = "importedData";
                int startRow = 1;
                int startCol = 1;

                #region // Adding of data to rc_data list //
                List<List<string>> rc_data = new List<List<string>>();

                // Adding Pipe data to rc_data
                foreach (Pipe p in pipes)
                {
                    rc_data.Add(new List<string>
                    {
                        p.Id.ToString(), p.Name, p.GetType().Name, p.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsString(),
                        "\"NA\"", "\"NA\"", "\"NA\"", "\"NA\"", "\"NA\"", p.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString(),
                        p.LookupParameter("Size").AsString() , p.LookupParameter("Length").AsValueString() , p.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString(),
                        IsVertical(p), GetPipeLevel(uidoc, p, "middle"), (12 * p.Diameter).ToString() + "\"", p.MEPSystem.Name,
                    });
                }

                //Adding Pipe Fitting Data to rc_data
                foreach (FamilyInstance inst in fittings)
                {
                    string mcaa_lbr_rate = inst?.LookupParameter("KDS_MCAA_LBR_RATE")?.AsValueString() ?? "\"NA\"";
                    string lbr_rate = inst?.LookupParameter("KDS_LBR_RATE")?.AsValueString() ?? "\"NA\"";
                    string hph = inst?.LookupParameter("KDS_HPH")?.AsString() ?? "\"NA\"";
                    string mfrList = inst?.LookupParameter("KDS_MfrList")?.AsValueString() ?? "\"NA\"";
                    string mfrPart = inst?.LookupParameter("KDS_MfrPart")?.AsString() ?? "\"NA\"";
                    string size = inst?.LookupParameter("Size")?.AsString() ?? "\"NA\"";

                    rc_data.Add(new List<string>
                    {
                        inst.Id.ToString(), inst.Name, inst.GetType().Name, inst.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsString(),
                        mcaa_lbr_rate, lbr_rate, hph, mfrList, mfrPart, inst.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString(), size,
                        "\"NA\"" , inst.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString(),
                        "\"NA\"", GetInstanceLevel(uidoc, inst), "\"NA\"", inst.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString(),
                    });
                }

                // Adding Plubing Fixture data to rc_data
                foreach (FamilyInstance f in fixtures)
                {
                    rc_data.Add(new List<string>
                    {
                        f.Id.ToString(), f.Name, f.GetType().Name, f.get_Parameter(BuiltInParameter.RBS_SYSTEM_CLASSIFICATION_PARAM).AsString(),
                        "\"NA\"", "\"NA\"", "\"NA\"", "\"NA\"", "\"NA\"", f.get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM).AsValueString(),
                        "\"NA\"" , "\"NA\"" , f.get_Parameter(BuiltInParameter.RBS_PIPING_SYSTEM_TYPE_PARAM).AsValueString(),
                        "\"NA\"", GetInstanceLevel(uidoc, f), "\"NA\"", f.get_Parameter(BuiltInParameter.RBS_SYSTEM_NAME_PARAM).AsString(),
                    });
                }
                #endregion

                CreateXL_macro(filePath, sheetName, startRow, startCol, rc_data, hdrList);

            } // end if (Yes) 
            #endregion

        } // end DialogBox Macro
        #endregion

        #region // GetPipes Function to collect all Pipe Elements //
        public static List<Pipe> GetPipes(UIDocument uidoc)
        {
            Document actvDoc = uidoc.Document;
            List<Pipe> elements = new List<Pipe>();

            FilteredElementCollector pipeCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(Pipe));

            foreach (Pipe p in pipeCollector)
            {
                elements.Add(p);
            }
            return elements;
        }
        #endregion

        #region // GetPipeFittings Function to collect all Pipe Fittings //
        public static List<FamilyInstance> GetPipeFittings(UIDocument uidoc)
        {
            Document actvDoc = uidoc.Document;
            List<FamilyInstance> elements = new List<FamilyInstance>();

            FilteredElementCollector fittingCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_PipeFitting);

            foreach (FamilyInstance inst in fittingCollector)
            {
                elements.Add(inst);
            }
            return elements;
        }
        #endregion

        #region // GetPlumbingFixtures Function to collect all Plumbing Fixtures //
        public static List<FamilyInstance> GetPlumbingFixtures(UIDocument uidoc)
        {
            Document actvDoc = uidoc.Document;
            List<FamilyInstance> elements = new List<FamilyInstance>();

            FilteredElementCollector fixtureCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_PlumbingFixtures);

            foreach (FamilyInstance f in fixtureCollector)
            {
                elements.Add(f);
            }
            return elements;
        }
        #endregion 

        #region // Get_DocLevel_strings Function to collect Names of Levels //
        public static List<string> Get_DocLevel_strings(UIDocument uidoc)
        {
            Document actvDoc = uidoc.Document;

            List<string> docLevels_Names_lst = new FilteredElementCollector(actvDoc)
                                                                     .OfClass(typeof(Level))
                                                                     .Select(e => e.Name)
                                                                     .Where(e => e != null)
                                                                     .ToList();
            return docLevels_Names_lst;
        }
        #endregion

        #region // GetPipeLevel Function to get Level Name of Current Pipe //
        public string GetPipeLevel(UIDocument uidoc, Pipe p, string startEndMiddle)
        {
            string lvlName = null;
            LocationCurve lc = p.Location as LocationCurve;
            Curve c = lc.Curve;
            double zLocStart = c.GetEndPoint(0).Z;
            double zLocEnd = c.GetEndPoint(1).Z;
            double zLocMiddle = (zLocStart + zLocEnd) / 2;

            Document actvDoc = uidoc.Document;

            List<Level> levels = new List<Level>();
            List<Level> sortedLevels = new List<Level>();

            FilteredElementCollector lvlCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(Level));

            foreach (Level level in lvlCollector)
            {
                levels.Add(level);
            }

            sortedLevels = levels.OrderBy(o => o.Elevation).ToList();

            #region // User wants Z-Coordinate of Start EndPoint for Pipe //

            if (startEndMiddle == "start" || startEndMiddle == "Start")
            {
                foreach (Level level in sortedLevels)
                {
                    if (zLocStart > level.Elevation)
                    {
                        lvlName = level.Name;
                    }
                }

                if (lvlName == null)
                {
                    lvlName = sortedLevels.First().Name;
                }
            }
            #endregion

            #region // User wants Z-Coordinate of End Endpoint for Pipe //
            if (startEndMiddle == "end" || startEndMiddle == "End")
            {
                foreach (Level level in sortedLevels)
                {
                    if (zLocEnd > level.Elevation)
                    {
                        lvlName = level.Name;
                    }
                }

                if (lvlName == null)
                {
                    lvlName = sortedLevels.First().Name;
                }
            }
            #endregion

            #region // User wants Z-Coordinate of Midpoint of Pipe //
            if (startEndMiddle == "middle" || startEndMiddle == "Middle")
            {
                foreach (Level level in sortedLevels)
                {
                    if (zLocMiddle > level.Elevation)
                    {
                        lvlName = level.Name;
                    }
                }

                if (lvlName == null)
                {
                    lvlName = sortedLevels.First().Name;
                }
            }
            #endregion

            else { TaskDialog.Show("error", "value for startEndMiddle is not \"start\", \"end\", or \"middle\""); }

            return lvlName;
        }
        #endregion

        #region // GetFittingLevel Function to get Level Name of Current Fitting //
        public string GetInstanceLevel(UIDocument uidoc, FamilyInstance inst)
        {
            Document actvDoc = uidoc.Document;

            string lvlName = null;
            LocationPoint lp = inst.Location as LocationPoint;
            double zLoc = lp.Point.Z;

            List<Level> levels = new List<Level>();
            List<Level> sortedLevels = new List<Level>();

            FilteredElementCollector lvlCollector = new FilteredElementCollector(actvDoc).OfClass(typeof(Level));

            foreach (Level level in lvlCollector)
            {
                levels.Add(level);
            }

            sortedLevels = levels.OrderBy(o => o.Elevation).ToList();

            foreach (Level level in sortedLevels)
            {
                if (zLoc > level.Elevation)
                {
                    lvlName = level.Name;
                }
            }

            if (lvlName == null)
            {
                lvlName = sortedLevels.First().Name;
            }

            return lvlName;
        }
        #endregion

        #region // IsVertical Function to determine whether a Pipe is vertical or not //
        public string IsVertical(Pipe p)
        {
            double tolerance = 0.01;
            double bottom_elv = p.get_Parameter(BuiltInParameter.RBS_PIPE_BOTTOM_ELEVATION).AsDouble();
            double top_elv = p.get_Parameter(BuiltInParameter.RBS_PIPE_TOP_ELEVATION).AsDouble();
            double length = p.LookupParameter("Length").AsDouble();
            double calcLength, lenDiff_perc;

            if (bottom_elv * top_elv * length != 0)
            {
                calcLength = Math.Abs(top_elv - bottom_elv);
                lenDiff_perc = Math.Abs((length - calcLength) / length);

                if (lenDiff_perc <= tolerance)
                {
                    return "Yes";
                }
                else
                {
                    return "No";
                }
            }
            else
            {
                return "NA";
            }
        }
        #endregion

        #region // CreateXL_macro Function to Handle Exportation of rc_data to excel //
        public static void CreateXL_macro(string filePath, string sheetName, int startRow, int startCol, List<List<string>> rc_data, List<string> hdrList)
        {
            #region // Check if FilePath Exists if not open filedialog foe user to choose a template file
            if (File.Exists(filePath))
            {
                //file exist
            }
            else
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = "C:\\";
                    openFileDialog.Filter = "Excel files (*.xltm)|*.xltm|All files (*.*)|*.*";
                    openFileDialog.FilterIndex = 2;
                    openFileDialog.RestoreDirectory = true;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //Get the path of specified file
                        filePath = openFileDialog.FileName;
                    }
                }
            }
            #endregion // Check if FilePath Exists if not open filedialog foe user to choose a template file

            //  Load the file into a byteArry to transfer to MemoryStream Later
            byte[] xltmDoc_byteArray = File.ReadAllBytes(filePath);
            using (MemoryStream xltmDoc_mem = new MemoryStream())
            {
                xltmDoc_mem.Write(xltmDoc_byteArray, 0, xltmDoc_byteArray.Length);

                using (SpreadsheetDocument xltmDoc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(xltmDoc_mem, true))
                {
                    xltmDoc.ChangeDocumentType(SpreadsheetDocumentType.MacroEnabledWorkbook);

                    #region  // Writing to Excel

                    WriteExcelFile(xltmDoc, sheetName, startRow, startCol, rc_data, hdrList);

                    #endregion  //  Writing to Excel

                    #region  // Set filepath to unique value

                    filePath = filePath.Replace("xltm", "xlsm");

                    // Check if filePath does not exist and increment name if it does.
                    filePath = AppendFileNumberIfExists(filePath, ".xlsm");

                    #endregion  // Set filepath to unique value

                }  // END using SpreadSheet

                #region  //  Change File Type and Save (XLTM to XLSM)
                // At this point, the memory stream contains the modified document.
                // We could write it back to a SharePoint document library or serve
                // it from a web server.
                // In this example, we serialize back to the file system to verify
                // that the code worked properly.
                using (FileStream fileStream = new FileStream(filePath, System.IO.FileMode.CreateNew))
                {
                    xltmDoc_mem.WriteTo(fileStream);

                }  // END using FileStream
                #endregion  //  Change File Type and Save (XLTM to XLSM)

            }  // END using MemoryStream
        }  // END of createXL
        #endregion

        #region // WriteExcelFile Function to write to data to excel file //
        private static void WriteExcelFile(SpreadsheetDocument xltmDoc, string sheetName, int startRow, int startCol, List<List<string>> rc_data, List<string> hdrList)
        {
            using (xltmDoc)
            {
                // Get the SharedStringTablePart. If it does not exist, create a new one.
                SharedStringTablePart shareStringPart;
                if (xltmDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = xltmDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = xltmDoc.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }

                #region  // Find desired worksheet by name or create one if it does not exist
                WorksheetPart worksheetPart = GetWorksheetPartByName(xltmDoc, sheetName);
                if (worksheetPart == null)
                {
                    // Insert a new worksheet of specified name .
                    worksheetPart = InsertWorksheet(xltmDoc.WorkbookPart, sheetName);
                }
                #endregion  // Find desired worksheet by name or create one if it does not exist

                int index;
                Cell cell;
                string cellValue = null;

                #region Loop thru rc_data to write data to excel file

                #region // Insert Header in row 0 //
                for (int r = 0; r < 1; r++)
                {
                    for (int c = 0; c < hdrList.Count(); c++)
                    {
                        #region Insert Values in Cells
                        // Insert the text into the SharedStringTablePart.
                        index = InsertSharedStringItem(hdrList[c], shareStringPart);

                        // Insert a cell into the worksheet.
                        cell = InsertCellInWorksheet(GetExcelColumnName(c + startCol), Convert.ToUInt32(r + startRow), worksheetPart);

                        // Set the value of cell.
                        cell.CellValue = new CellValue(index.ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        cellValue += "\n" + "column: " + GetExcelColumnName(c + startCol) + " : row : " + " " + Convert.ToUInt32(r + startRow) + " : " + index.ToString();

                        #endregion Insert Values in Cells
                    }
                }
                #endregion

                #region //Insert rc_data starting in row 1 //
                for (int r = 1; r < rc_data.Count; r++)
                {
                    for (int c = 0; c < rc_data[r].Count(); c++)
                    {
                        #region Insert Values in Cells
                        // Insert the text into the SharedStringTablePart.
                        index = InsertSharedStringItem(rc_data[r][c], shareStringPart);


                        // Insert a cell into the worksheet.
                        cell = InsertCellInWorksheet(GetExcelColumnName(c + startCol), Convert.ToUInt32(r + startRow), worksheetPart);

                        // Set the value of cell.
                        cell.CellValue = new CellValue(index.ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                        cellValue += "\n" + "column: " + GetExcelColumnName(c + startCol) + " : row : " + " " + Convert.ToUInt32(r + startRow) + " : " + index.ToString();

                        #endregion Insert Values in Cells
                    }
                }
                #endregion Loop thru list of Lists

                #endregion

                // Save the new worksheet.
                worksheetPart.Worksheet.Save();
            }  // END using xltmDoc
        }
        #endregion //  write to excel

        #region // InsertSharedStringItem Function to insert text into SharedStringTablePart //
        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }
            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }
        #endregion

        #region // InsertWorksheet Function to add new workseet to excel workbook //
        // Given a WorkbookPart, inserts a new worksheet.
        private static WorksheetPart InsertWorksheet(WorkbookPart workbookPart, string sheetName)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }
        #endregion

        #region // InsertCellinWorksheet Function to insert a cell into worksheet //
        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }
        #endregion

        #region // GetWorksheetPartByName Function //
        private static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument xltmDoc, string sheetName)
        {
            IEnumerable<Sheet> sheets =
               xltmDoc.WorkbookPart.Workbook.GetFirstChild<Sheets>().
               Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets?.Count() == 0)
            {
                // The specified worksheet does not exist.

                return null;
            }

            string relationshipId = sheets?.First().Id.Value;

            WorksheetPart worksheetPart = (WorksheetPart)xltmDoc.WorkbookPart.GetPartById(relationshipId);

            return worksheetPart;
        }
        #endregion

        #region // GetExcelColumnName Function to convert column number to letter(s) //
        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }
            return columnName;
        }
        #endregion

        #region // AppendFileNumberIfExists Function to Append File name //
        // <summary>
        /// A function to add an incremented number at the end of a file name if a file already exists. 
        /// </summary>
        /// <param name="file">The file. This should be the complete path.</param>
        /// <param name="ext">This can be empty.</param>
        /// <returns>An incremented file name. </returns>
        private static string AppendFileNumberIfExists(string file, string ext)
        {
            // This had a VB tidbit that helped to get this started. 
            // http://www.codeproject.com/Questions/212217/increment-filename-if-file-exists-using-csharp

            // If the file exists, then do stuff. Otherwise, we just return the original file name.
            if (File.Exists(file))
            {
                string folderPath = Path.GetDirectoryName(file); // The path to the file. No sense in dealing with this unecessarily. 
                string fileName = Path.GetFileNameWithoutExtension(file); // The file name with no extension. 
                string extension = string.Empty; // The file extension. 
                                                 // This lets us pass in an empty string for the file extension if required. i.e. It just makes this function a bit more versatile. 
                if (ext == string.Empty)
                {
                    extension = Path.GetExtension(file);
                }
                else
                {
                    extension = ext;
                }

                // at this point, find out if the fileName ends in a number, then get that number.
                int fileNumber = 0; // This stores the number as a number for us. 
                                    // need a regex here - \(([0-9]+)\)$
                Regex r = new Regex(@"\(([0-9]+)\)$"); // This matches the pattern we are using, i.e. ~(#).ext
                Match m = r.Match(fileName); // We pass in the file name with no extension.
                string addSpace = " "; // We'll add a space when we don't have our pattern in order to pad the pattern.
                if (m.Success)
                {
                    addSpace = string.Empty; // We have the pattern, so we don't add a space - it has already been added. 
                    string s = m.Groups[1].Captures[0].Value; // This is the single capture that we are looking for. Stored as a string.
                                                              // set fileNumber to the new number.
                    fileNumber = int.Parse(s); // Convert the number to an int.
                                               // remove the numbering from the string as we're constructing it again below.
                    fileName = fileName.Replace("(" + s + ")", "");
                }

                // Start looping. 
                do
                {
                    fileNumber += 1; // Increment the file number that we have above. 
                    file = Path.Combine(folderPath, // Combine it all.
                                            string.Format("{0}{3}({1}){2}", // The pattern to combine.
                                                                      fileName,         // The file name with no extension. 
                                                                      fileNumber,       // The file number.
                                                                      extension,        // The file extension.
                                                                      addSpace));       // A space if needed to pad the initial ~(#).ext pattern.
                }
                while (File.Exists(file)); // As long as the file name exists, keep looping. 
            }
            return file;
        }
        #endregion
    }
}
