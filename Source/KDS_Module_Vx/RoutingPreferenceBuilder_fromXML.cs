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
using Autodesk.Revit.DB.Plumbing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;


namespace Revit.SDK.Samples.RoutingPreferenceTools.CS
{
    /// <summary>
    /// Class to read and write XML and routing preference data
    /// </summary>
    public class RoutingPreferenceBuilder
    {
        #region Data
        private IEnumerable<Segment> m_segments;
        private IEnumerable<FamilySymbol> m_fittings;
        private IEnumerable<Material> m_materials;
        private IEnumerable<PipeScheduleType> m_pipeSchedules;
        private IEnumerable<PipeType> m_pipeTypes;
        private Autodesk.Revit.DB.Document m_document;
        #endregion

        #region Public interface
        /// <summary>
        /// Create an instance of the class and initialize lists of all segments, fittings, materials, schedules, and pipe types in the document.
        /// </summary>
        public RoutingPreferenceBuilder(Document document)
        {
            m_document = document;
            m_segments = GetAllPipeSegments(m_document);
            m_fittings = GetAllFittings(m_document);
            m_materials = GetAllMaterials(m_document);
            m_pipeSchedules = GetAllPipeScheduleTypes(m_document);
            m_pipeTypes = GetAllPipeTypes(m_document);

        }
        /// <summary>
        /// Reads data from an Xml source and loads pipe fitting families, creates segments, sizes, schedules, and routing preference rules from the xml data.
        /// </summary>
        /// <param name="xDoc">The Xml data source to read from</param>
        public void ParseAllPipingPoliciesFromXml(XDocument xDoc)
        {
            //Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "I am in ParseAllPipingPoliciesFromXml. PipeTypes.Count = " + m_pipeTypes.Count());
            if (m_pipeTypes.Count() == 0)
                throw new RoutingPreferenceDataException("No pipe pipes defined in this project.  At least one must be defined.");


            FormatOptions formatOptionPipeSize = m_document.GetUnits().GetFormatOptions(UnitType.UT_PipeSize);

            string docPipeSizeUnit = formatOptionPipeSize.DisplayUnits.ToString();
            string xmlPipeSizeUnit = xDoc.Root.Attribute("pipeSizeUnits").Value;
            if (docPipeSizeUnit != xmlPipeSizeUnit)
                throw new RoutingPreferenceDataException("Units from XML do not match current pipe size units.");
            //Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "I am in ParseAllPipingPoliciesFromXml. xmlPipeSizeUnit = " + xmlPipeSizeUnit);

            FormatOptions formatOptionRoughness = m_document.GetUnits().GetFormatOptions(UnitType.UT_Piping_Roughness);

            string docRoughnessUnit = formatOptionRoughness.DisplayUnits.ToString();
            string xmlRoughnessUnit = xDoc.Root.Attribute("pipeRoughnessUnits").Value;
            //Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "I am in ParseAllPipingPoliciesFromXml. docRoughnessUnit = " + docRoughnessUnit);
            //Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "I am in ParseAllPipingPoliciesFromXml. xmlRoughnessUnit = " + xmlRoughnessUnit);

            if (docRoughnessUnit != xmlRoughnessUnit)
                throw new RoutingPreferenceDataException("Units from XML do not match current pipe roughness units.");
            //if (m_document.IsModifiable) { Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "I am in ParseAllPipingPoliciesFromXml. Document IS modifiable"); }
            //if (m_document.IsReadOnly) { Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "I am in ParseAllPipingPoliciesFromXml. Document is IsReadOnly"); }
            //else { Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "I am in ParseAllPipingPoliciesFromXml. Document is NOT ReadOnly"); }
            //if (m_document.IsReadOnlyFile) { Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "I am in ParseAllPipingPoliciesFromXml. Document is IsReadOnlyFile"); }
            //else { Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "I am in ParseAllPipingPoliciesFromXml. Document is NOT ReadOnlyFile"); }


            Transaction loadFamilies = new Transaction(m_document, "loadFamilies");

            //Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "I am in ParseAllPipingPoliciesFromXml. I will Start Transaction loadFamilies. loadFamilies GetStatus:  " + loadFamilies.GetStatus().ToString());

            if (loadFamilies.Start().Equals(TransactionStatus.Started))
            {
                Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "I am in ParseAllPipingPoliciesFromXml. Just Started Transaction Loading Families = "); // + loadFamilies.GetStatus().ToString());
            }
            else
            {
                Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "I am in ParseAllPipingPoliciesFromXml. Could not Start LoadFamilies:"); // + loadFamilies.GetStatus().ToString());
            }
            IEnumerable<XElement> families = xDoc.Root.Elements("Family");
            FindFolderUtility findFolderUtility = new FindFolderUtility(m_document.Application);
            //Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "I am in ParseAllPipingPoliciesFromXml. families = " + families.ToList());
            foreach (XElement xfamily in families)
                try
                {
                    ParseFamilyFromXml(xfamily, findFolderUtility);  //Load families.
                }
                catch (Exception ex)
                {
                    loadFamilies.RollBack();
                    throw ex;
                }
            loadFamilies.Commit();


            Transaction addPipeTypes = new Transaction(m_document, "Add PipeTypes");
            addPipeTypes.Start();
            IEnumerable<XElement> pipeTypes = xDoc.Root.Elements("PipeType");
            foreach (XElement xpipeType in pipeTypes)
                try
                {
                    ParsePipeTypeFromXml(xpipeType);  //Define new pipe types.
                }
                catch (Exception ex)
                {
                    addPipeTypes.RollBack();
                    throw ex;
                }
            addPipeTypes.Commit();

            Transaction addPipeSchedules = new Transaction(m_document, "Add Pipe Schedule Types");
            addPipeSchedules.Start();
            IEnumerable<XElement> pipeScheduleTypes = xDoc.Root.Elements("PipeScheduleType");
            foreach (XElement xpipeScheduleType in pipeScheduleTypes)
                try
                {
                    ParsePipeScheduleTypeFromXml(xpipeScheduleType);  //Define new pipe schedule types.
                }
                catch (Exception ex)
                {
                    addPipeSchedules.RollBack();
                    throw ex;
                }
            addPipeSchedules.Commit();

            //The code above have added some new pipe types, schedules, or fittings, so update the lists of all of these.
            UpdatePipeTypesList();
            UpdatePipeTypeSchedulesList();
            UpdateFittingsList();

            Transaction addPipeSegments = new Transaction(m_document, "Add Pipe Segments");
            addPipeSchedules.Start();
            IEnumerable<XElement> pipeSegments = xDoc.Root.Elements("PipeSegment");  //Define new segments.
            foreach (XElement xpipeSegment in pipeSegments)
                try
                {
                    ParsePipeSegmentFromXML(xpipeSegment);
                }
                catch (Exception ex)
                {
                    addPipeSchedules.RollBack();
                    throw ex;
                }
            addPipeSchedules.Commit();

            UpdateSegmentsList();  //More segments may have been created, so update the segment list.


            //Now that all of the various types that routing preferences use have been created or loaded, add all the routing preferences.
            Transaction addRoutingPreferences = new Transaction(m_document, "Add Routing Preferences");
            addRoutingPreferences.Start();
            IEnumerable<XElement> routingPreferenceManagers = xDoc.Root.Elements("RoutingPreferenceManager");
            foreach (XElement xroutingPreferenceManager in routingPreferenceManagers)
                try
                {
                    ParseRoutingPreferenceManagerFromXML(xroutingPreferenceManager);
                }
                catch (Exception ex)
                {
                    addRoutingPreferences.RollBack();
                    throw ex;
                }
            addRoutingPreferences.Commit();
            Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilder", "I am Done with  ParseAllPipingPoliciesFromXml ");

        }

n

        #region XML parsing and generation
        /// <summary>
        /// Load a family from xml
        /// </summary>
        /// <param name="familyXElement"></param>
        /// <param name="findFolderUtility"></param>
        private void ParseFamilyFromXml(XElement familyXElement, FindFolderUtility findFolderUtility)
        {
            Autodesk.Revit.UI.TaskDialog.Show("RoutingPreferenceBuilderUtility", "Starting ParseFamilyFromXml");

            XAttribute xafilename = familyXElement.Attribute(XName.Get("filename"));
            string familyPath = xafilename.Value;
            if (!System.IO.File.Exists(familyPath))
            {
                string filename = System.IO.Path.GetFileName(familyPath);
                familyPath = findFolderUtility.FindFileFolder(filename);
                if (!System.IO.File.Exists(familyPath))
                    throw new RoutingPreferenceDataException("Cannot find family file: " + xafilename.Value);
            }


            if (string.Compare(System.IO.Path.GetExtension(familyPath), ".rfa", true) != 0)
                throw new RoutingPreferenceDataException(familyPath + " is not a family file.");

            try
            {
                if (!m_document.LoadFamily(familyPath))
                    return;  //returns false if already loaded.
            }
            catch (System.Exception ex)
            {
                throw new RoutingPreferenceDataException("Cannot load family: " + xafilename.Value + ": " + ex.ToString());
            }

        }


        /// <summary>
        /// Greate a PipeType from xml
        /// </summary>
        /// <param name="pipetypeXElement"></param>
        private void ParsePipeTypeFromXml(XElement pipetypeXElement)
        {
            XAttribute xaName = pipetypeXElement.Attribute(XName.Get("name"));

            ElementId pipeTypeId = GetPipeTypeByName(xaName.Value);

            if (pipeTypeId == ElementId.InvalidElementId)  //If the pipe type does not exist, create it.
            {
                PipeType newPipeType = m_pipeTypes.First().Duplicate(xaName.Value) as PipeType;
                ClearRoutingPreferenceRules(newPipeType);
            }

        }

        /// <summary>
        /// Clear all routing preferences in a PipeType
        /// </summary>
        /// <param name="pipeType"></param>
        private static void ClearRoutingPreferenceRules(PipeType pipeType)
        {
            foreach (RoutingPreferenceRuleGroupType group in System.Enum.GetValues(typeof(RoutingPreferenceRuleGroupType)))
            {
                int ruleCount = pipeType.RoutingPreferenceManager.GetNumberOfRules(group);
                for (int index = 0; index != ruleCount; ++index)
                {
                    pipeType.RoutingPreferenceManager.RemoveRule(group, 0);
                }
            }
        }


        private void ParsePipeScheduleTypeFromXml(XElement pipeScheduleTypeXElement)
        {
            XAttribute xaName = pipeScheduleTypeXElement.Attribute(XName.Get("name"));
            ElementId pipeScheduleTypeId = GetPipeScheduleTypeByName(xaName.Value);
            if (pipeScheduleTypeId == ElementId.InvalidElementId)  //If the pipe schedule type does not exist, create it.
                m_pipeSchedules.First().Duplicate(xaName.Value);
        }


        /// <summary>
        /// Create a PipeSegment from XML
        /// </summary>
        /// <param name="segmentXElement"></param>
        private void ParsePipeSegmentFromXML(XElement segmentXElement)
        {
            XAttribute xaMaterial = segmentXElement.Attribute(XName.Get("materialName"));
            XAttribute xaSchedule = segmentXElement.Attribute(XName.Get("pipeScheduleTypeName"));
            XAttribute xaRoughness = segmentXElement.Attribute(XName.Get("roughness"));

            ElementId materialId = GetMaterialByName(xaMaterial.Value);  //There is nothing in the xml schema for creating new materials -- any material specified must already exist in the document.
            if (materialId == ElementId.InvalidElementId)
            {
                throw new RoutingPreferenceDataException("Cannot find Material: " + xaMaterial.Value + " in: " + segmentXElement.ToString());
            }
            ElementId scheduleId = GetPipeScheduleTypeByName(xaSchedule.Value);

            double roughness;
            bool r1 = double.TryParse(xaRoughness.Value, out roughness);

            if (!r1)
                throw new RoutingPreferenceDataException("Invalid roughness value: " + xaRoughness.Value + " in: " + segmentXElement.ToString());

            if (roughness <= 0)
                throw new RoutingPreferenceDataException("Invalid roughness value: " + xaRoughness.Value + " in: " + segmentXElement.ToString());

            if (scheduleId == ElementId.InvalidElementId)
            {
                throw new RoutingPreferenceDataException("Cannot find Schedule: " + xaSchedule.Value + " in: " + segmentXElement.ToString());  //we will not create new schedules.
            }

            ElementId existingPipeSegmentId = GetSegmentByIds(materialId, scheduleId);
            if (existingPipeSegmentId != ElementId.InvalidElementId)
                return;   //Segment found, no need to create.

            ICollection<MEPSize> sizes = new List<MEPSize>();
            foreach (XNode sizeNode in segmentXElement.Nodes())
            {
                if (sizeNode is XElement)
                {
                    MEPSize newSize = ParseMEPSizeFromXml(sizeNode as XElement, m_document);
                    sizes.Add(newSize);
                }
            }
            PipeSegment pipeSegment = PipeSegment.Create(m_document, materialId, scheduleId, sizes);
            pipeSegment.Roughness = Utility.Convert.ConvertValueToFeet(roughness, m_document);

            return;
        }


        /// <summary>
        /// Create an MEPSize from Xml
        /// </summary>
        /// <param name="sizeXElement"></param>
        /// <param name="document"></param>
        /// <returns></returns>
        private static MEPSize ParseMEPSizeFromXml(XElement sizeXElement, Autodesk.Revit.DB.Document document)
        {
            XAttribute xaNominal = sizeXElement.Attribute(XName.Get("nominalDiameter"));
            XAttribute xaInner = sizeXElement.Attribute(XName.Get("innerDiameter"));
            XAttribute xaOuter = sizeXElement.Attribute(XName.Get("outerDiameter"));
            XAttribute xaUsedInSizeLists = sizeXElement.Attribute(XName.Get("usedInSizeLists"));
            XAttribute xaUsedInSizing = sizeXElement.Attribute(XName.Get("usedInSizing"));

            double nominal, inner, outer;
            bool usedInSizeLists, usedInSizing;
            bool r1 = double.TryParse(xaNominal.Value, out nominal);
            bool r2 = double.TryParse(xaInner.Value, out inner);
            bool r3 = double.TryParse(xaOuter.Value, out outer);
            bool r4 = bool.TryParse(xaUsedInSizeLists.Value, out usedInSizeLists);
            bool r5 = bool.TryParse(xaUsedInSizing.Value, out usedInSizing);

            if (!r1 || !r2 || !r3 || !r4 || !r5)
                throw new RoutingPreferenceDataException("Cannot parse MEPSize attributes:" + xaNominal.Value + ", " + xaInner.Value + ", " + xaOuter.Value + ", " + xaUsedInSizeLists.Value + ", " + xaUsedInSizing.Value);

            MEPSize newSize = null;

            try
            {

                newSize = new MEPSize(Utility.Convert.ConvertValueToFeet(nominal, document), Utility.Convert.ConvertValueToFeet(inner, document), Utility.Convert.ConvertValueToFeet(outer, document), usedInSizeLists, usedInSizing);
            }

            catch (Exception)
            {
                throw new RoutingPreferenceDataException("Invalid MEPSize values: " + nominal.ToString() + ", " + inner.ToString() + ", " + outer.ToString());
            }
            return newSize;

        }


        /// <summary>
        /// Populate a routing preference manager from Xml
        /// </summary>
        /// <param name="routingPreferenceManagerXElement"></param>
        private void ParseRoutingPreferenceManagerFromXML(XElement routingPreferenceManagerXElement)
        {

            XAttribute xaPipeTypeName = routingPreferenceManagerXElement.Attribute(XName.Get("pipeTypeName"));
            XAttribute xaPreferredJunctionType = routingPreferenceManagerXElement.Attribute(XName.Get("preferredJunctionType"));

            PreferredJunctionType preferredJunctionType;
            bool r1 = Enum.TryParse<PreferredJunctionType>(xaPreferredJunctionType.Value, out preferredJunctionType);

            if (!r1)
                throw new RoutingPreferenceDataException("Invalid Preferred Junction Type in: " + routingPreferenceManagerXElement.ToString());

            ElementId pipeTypeId = GetPipeTypeByName(xaPipeTypeName.Value);
            if (pipeTypeId == ElementId.InvalidElementId)
                throw new RoutingPreferenceDataException("Could not find pipe type element in: " + routingPreferenceManagerXElement.ToString());

            PipeType pipeType = m_document.GetElement(pipeTypeId) as PipeType;

            RoutingPreferenceManager routingPreferenceManager = pipeType.RoutingPreferenceManager;
            routingPreferenceManager.PreferredJunctionType = preferredJunctionType;

            foreach (XNode xRule in routingPreferenceManagerXElement.Nodes())
            {
                if (xRule is XElement)
                {
                    RoutingPreferenceRuleGroupType groupType;
                    RoutingPreferenceRule rule = ParseRoutingPreferenceRuleFromXML(xRule as XElement, out groupType);
                    routingPreferenceManager.AddRule(groupType, rule);
                    //routingPreferenceManager.AddRule(groupType, rule, ruleIndex);
                }
            }

        }
        /// <summary>

        /// <summary>
        /// Create a RoutingPreferenceRule from Xml
        /// </summary>
        /// <param name="ruleXElement"></param>
        /// <param name="groupType"></param>
        /// <returns></returns>
        private RoutingPreferenceRule ParseRoutingPreferenceRuleFromXML(XElement ruleXElement, out RoutingPreferenceRuleGroupType groupType)
        {

            XAttribute xaDescription = null;
            XAttribute xaPartName = null;
            XAttribute xaMinSize = null;
            XAttribute xaMaxSize = null;
            XAttribute xaGroup = null;
            XAttribute xruleIndex = null;

            xaDescription = ruleXElement.Attribute(XName.Get("description"));
            xaPartName = ruleXElement.Attribute(XName.Get("partName"));
            xaGroup = ruleXElement.Attribute(XName.Get("ruleGroup"));
            xaMinSize = ruleXElement.Attribute(XName.Get("minimumSize"));
            xruleIndex = ruleXElement.Attribute(XName.Get("ruleIndex"));

            ElementId partId;

            bool r3 = Enum.TryParse<RoutingPreferenceRuleGroupType>(xaGroup.Value, out groupType);
            if (!r3)
                throw new RoutingPreferenceDataException("Could not parse rule group type: " + xaGroup.Value);

            string description = xaDescription.Value;

            if (groupType == RoutingPreferenceRuleGroupType.Segments)
                partId = GetSegmentByName(xaPartName.Value);
            else
                partId = GetFittingByName(xaPartName.Value);

            if (partId == ElementId.InvalidElementId)
                throw new RoutingPreferenceDataException("Could not find MEP Part: " + xaPartName.Value + ".  Is this the correct family name, and is the correct family loaded?");

            RoutingPreferenceRule rule = new RoutingPreferenceRule(partId, description);


            PrimarySizeCriterion sizeCriterion;
            if (string.Compare(xaMinSize.Value, "All", true) == 0)  //If "All" or "None" are specified, set min and max values to documented "Max" values.
            {
                sizeCriterion = PrimarySizeCriterion.All();
            }
            else if (string.Compare(xaMinSize.Value, "None", true) == 0)
            {
                sizeCriterion = PrimarySizeCriterion.None();
            }
            else  // "maximumSize" attribute is only needed if not specifying "All" or "None."
            {
                try
                {
                    xaMaxSize = ruleXElement.Attribute(XName.Get("maximumSize"));
                }
                catch (System.Exception)
                {
                    throw new RoutingPreferenceDataException("Cannot get maximumSize attribute in: " + ruleXElement.ToString());
                }
                double min, max;
                bool r1 = double.TryParse(xaMinSize.Value, out min);
                bool r2 = double.TryParse(xaMaxSize.Value, out max);
                if (!r1 || !r2)
                    throw new RoutingPreferenceDataException("Could not parse size values: " + xaMinSize.Value + ", " + xaMaxSize.Value);
                if (min > max)
                    throw new RoutingPreferenceDataException("Invalid size range.");

                min = Utility.Convert.ConvertValueToFeet(min, m_document);
                max = Utility.Convert.ConvertValueToFeet(max, m_document);
                sizeCriterion = new PrimarySizeCriterion(min, max);
            }

            rule.AddCriterion(sizeCriterion);

            return rule;

        }



        #region Accessors and finders
        /// <summary>
        /// Get PipeScheduleTypeName by Id
        /// </summary>
        /// <param name="pipescheduleTypeId"></param>
        /// <returns></returns>
        private string GetPipeScheduleTypeNamebyId(ElementId pipescheduleTypeId)
        {
            return m_document.GetElement(pipescheduleTypeId).Name;
        }

        /// <summary>
        /// Get material name by Id
        /// </summary>
        /// <param name="materialId"></param>
        /// <returns></returns>
        private string GetMaterialNameById(ElementId materialId)
        {
            return m_document.GetElement(materialId).Name;
        }

        /// <summary>
        /// Get segment name by Id
        /// </summary>
        /// <param name="segmentId"></param>
        /// <returns></returns>
        private string GetSegmentNameById(ElementId segmentId)
        {
            return m_document.GetElement(segmentId).Name;
        }

        /// <summary>
        /// Get fitting name by Id
        /// </summary>
        /// <param name="fittingId"></param>
        /// <returns></returns>
        private string GetFittingNameById(ElementId fittingId)
        {
            FamilySymbol fs = m_document.GetElement(fittingId) as FamilySymbol;
            return fs.Family.Name + " " + fs.Name;
        }

        /// <summary>
        /// Get segment by Ids
        /// </summary>
        /// <param name="materialId"></param>
        /// <param name="pipeScheduleTypeId"></param>
        /// <returns></returns>
        private ElementId GetSegmentByIds(ElementId materialId, ElementId pipeScheduleTypeId)
        {
            if ((materialId == ElementId.InvalidElementId) || (pipeScheduleTypeId == ElementId.InvalidElementId))
                return ElementId.InvalidElementId;

            Element material = m_document.GetElement(materialId);
            Element pipeScheduleType = m_document.GetElement(pipeScheduleTypeId);
            string segmentName = material.Name + " - " + pipeScheduleType.Name;
            return GetSegmentByName(segmentName);

        }

        /// <summary>
        /// Get pipe type name by Id
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        private string GetPipeTypeNameById(ElementId id)
        {
            return m_document.GetElement(id).Name;
        }

        /// <summary>
        /// Get segment by name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private ElementId GetSegmentByName(string name)
        {
            foreach (Segment segment in m_segments)
                if (segment.Name == name)
                    return segment.Id;
            return ElementId.InvalidElementId;
        }

        /// <summary>
        /// Get fitting by name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private ElementId GetFittingByName(string name)
        {
            foreach (FamilySymbol fitting in m_fittings)
                if ((fitting.Family.Name + " " + fitting.Name) == name)

                    return fitting.Id;
            return ElementId.InvalidElementId;
        }

        /// <summary>
        /// Get material by name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private ElementId GetMaterialByName(string name)
        {
            foreach (Material material in m_materials)
                if (material.Name == name)
                    return material.Id;
            return ElementId.InvalidElementId;
        }

        /// <summary>
        /// Get pipe schedule type by name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private ElementId GetPipeScheduleTypeByName(string name)
        {
            foreach (PipeScheduleType pipeScheduleType in m_pipeSchedules)
                if (pipeScheduleType.Name == name)
                    return pipeScheduleType.Id;
            return ElementId.InvalidElementId;

        }

        /// <summary>
        /// Get pipe type by name
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private ElementId GetPipeTypeByName(string name)
        {
            foreach (PipeType pipeType in m_pipeTypes)
                if (pipeType.Name == name)
                    return pipeType.Id;
            return ElementId.InvalidElementId;

        }

        /// <summary>
        /// Update fittings list
        /// </summary>
        private void UpdateFittingsList()
        {
            m_fittings = GetAllFittings(m_document);
        }

        /// <summary>
        /// Update segments list
        /// </summary>
        private void UpdateSegmentsList()
        {
            m_segments = GetAllPipeSegments(m_document);
        }

        /// <summary>
        /// Update pipe types list
        /// </summary>
        private void UpdatePipeTypesList()
        {
            m_pipeTypes = GetAllPipeTypes(m_document);
        }

        /// <summary>
        /// Update pipe type schedules list
        /// </summary>
        private void UpdatePipeTypeSchedulesList()
        {
            m_pipeSchedules = GetAllPipeScheduleTypes(m_document);
        }

        /// <summary>
        /// Update materials list
        /// </summary>
        private void UpdateMaterialsList()
        {
            m_materials = GetAllMaterials(m_document);
        }

        /// <summary>
        /// Get all pipe segments
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public IEnumerable<PipeSegment> GetAllPipeSegments(Document document)
        {
            FilteredElementCollector fec = new FilteredElementCollector(document);
            fec.OfClass(typeof(PipeSegment));
            IEnumerable<PipeSegment> segments = fec.ToElements().Cast<PipeSegment>();
            return segments;
        }

        /// <summary>
        /// Get all fittings
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public IEnumerable<FamilySymbol> GetAllFittings(Document document)
        {
            FilteredElementCollector fec = new FilteredElementCollector(document);
            fec.OfClass(typeof(FamilySymbol));
            fec.OfCategory(BuiltInCategory.OST_PipeFitting);
            IEnumerable<FamilySymbol> fittings = fec.ToElements().Cast<FamilySymbol>();
            return fittings;
        }

        /// <summary>
        /// Get all materials
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        private IEnumerable<Material> GetAllMaterials(Document document)
        {
            FilteredElementCollector fec = new FilteredElementCollector(document);
            fec.OfClass(typeof(Material));
            IEnumerable<Material> materials = fec.ToElements().Cast<Material>();
            return materials;
        }

        /// <summary>
        /// Get all pipe schedule types
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public IEnumerable<PipeScheduleType> GetAllPipeScheduleTypes(Document document)
        {
            FilteredElementCollector fec = new FilteredElementCollector(document);
            fec.OfClass(typeof(Autodesk.Revit.DB.Plumbing.PipeScheduleType));
            IEnumerable<Autodesk.Revit.DB.Plumbing.PipeScheduleType> pipeScheduleTypes = fec.ToElements().Cast<Autodesk.Revit.DB.Plumbing.PipeScheduleType>();
            return pipeScheduleTypes;
        }

        /// <summary>
        /// Get all pipe types
        /// </summary>
        /// <param name="document"></param>
        /// <returns></returns>
        public IEnumerable<PipeType> GetAllPipeTypes(Document document)
        {
            ElementClassFilter ecf = new ElementClassFilter(typeof(PipeType));

            FilteredElementCollector fec = new FilteredElementCollector(document);
            fec.WherePasses(ecf);
            IEnumerable<PipeType> pipeTypes = fec.ToElements().Cast<PipeType>();
            return pipeTypes;
        }


        #endregion


    }
}
