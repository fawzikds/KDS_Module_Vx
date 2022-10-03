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
        /// Reads pipe fitting family, segment, size, schedule, and routing preference data from a document and summarizes it in Xml.
        /// </summary>
        /// <returns>An XDocument containing an Xml summary of routing preference information</returns>
        public XDocument CreateXmlFromAllPipingPolicies(ref bool pathsNotFound)
        {
            //To export the full path name of all .rfa family files, use the FindFolderUtility class.
            FindFolderUtility findFolderUtility = new FindFolderUtility(m_document.Application);

            XDocument routingPreferenceBuilderDoc = new XDocument();
            //XElement xroot = new XElement(XName.Get("RoutingPreferenceBuilder"));
            XElement xroot = new XElement(XName.Get("RoutingPreferenceBuilder_ruleIndex"));

            FormatOptions formatOptionPipeSize = m_document.GetUnits().GetFormatOptions(UnitType.UT_PipeSize);
            string unitStringPipeSize = formatOptionPipeSize.DisplayUnits.ToString();
            xroot.Add(new XAttribute(XName.Get("pipeSizeUnits"), unitStringPipeSize));

            FormatOptions formatOptionRoughness = m_document.GetUnits().GetFormatOptions(UnitType.UT_Piping_Roughness);
            string unitStringRoughness = formatOptionRoughness.DisplayUnits.ToString();
            xroot.Add(new XAttribute(XName.Get("pipeRoughnessUnits"), unitStringRoughness));

            foreach (FamilySymbol familySymbol in this.m_fittings)
            {
                xroot.Add(CreateXmlFromFamily(familySymbol, findFolderUtility, ref pathsNotFound));
            }

            foreach (PipeType pipeType in m_pipeTypes)
            {
                xroot.Add(CreateXmlFromPipeType(pipeType));
            }

            foreach (PipeScheduleType pipeScheduleType in m_pipeSchedules)
            {
                xroot.Add(CreateXmlFromPipeScheduleType(pipeScheduleType));
            }

            foreach (PipeSegment pipeSegment in m_segments)
            {
                xroot.Add(CreateXmlFromPipeSegment(pipeSegment));
            }

            foreach (PipeType pipeType in m_pipeTypes)
            {
                xroot.Add(CreateXmlFromRoutingPreferenceManager(pipeType.RoutingPreferenceManager));
            }

            routingPreferenceBuilderDoc.Add(xroot);
            return routingPreferenceBuilderDoc;
        }
        #endregion


        /// <summary>
        /// Create xml from a family
        /// </summary>
        /// <param name="pipeFitting"></param>
        /// <param name="findFolderUtility"></param>
        /// <param name="pathNotFound"></param>
        /// <returns></returns>
        private static XElement CreateXmlFromFamily(FamilySymbol pipeFitting, FindFolderUtility findFolderUtility, ref bool pathNotFound)
        {
            //Try to find the path of the .rfa file.
            string path = findFolderUtility.FindFileFolder(pipeFitting.Family.Name + ".rfa");
            string pathToWrite;
            if (path == "")
            {
                pathNotFound = true;
                pathToWrite = pipeFitting.Family.Name + ".rfa";
            }
            else
                pathToWrite = path;

            XElement xFamilySymbol = new XElement(XName.Get("Family"));
            xFamilySymbol.Add(new XAttribute(XName.Get("filename"), pathToWrite));
            return xFamilySymbol;
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

        /// <summary>
        /// Create Xml from a PipeType
        /// </summary>
        /// <param name="pipeType"></param>
        /// <returns></returns>
        private static XElement CreateXmlFromPipeType(PipeType pipeType)
        {
            XElement xPipeType = new XElement(XName.Get("PipeType"));
            xPipeType.Add(new XAttribute(XName.Get("name"), pipeType.Name));
            return xPipeType;
        }


        /// <summary>
        /// Create Xml from a PipeScheduleType
        /// </summary>
        /// <param name="pipeScheduleType"></param>
        /// <returns></returns>
        private static XElement CreateXmlFromPipeScheduleType(PipeScheduleType pipeScheduleType)
        {
            XElement xPipeSchedule = new XElement(XName.Get("PipeScheduleType"));
            xPipeSchedule.Add(new XAttribute(XName.Get("name"), pipeScheduleType.Name));
            return xPipeSchedule;
        }


        /// <summary>
        /// Create Xml from a PipeSegment
        /// </summary>
        /// <param name="pipeSegment"></param>
        /// <returns></returns>
        private XElement CreateXmlFromPipeSegment(PipeSegment pipeSegment)
        {
            XElement xPipeSegment = new XElement(XName.Get("PipeSegment"));
            string segmentName = pipeSegment.Name;

            xPipeSegment.Add(new XAttribute(XName.Get("pipeScheduleTypeName"), GetPipeScheduleTypeNamebyId(pipeSegment.ScheduleTypeId)));
            xPipeSegment.Add(new XAttribute(XName.Get("materialName"), GetMaterialNameById(pipeSegment.MaterialId)));

            double roughnessInDocumentUnits = Utility.Convert.ConvertValueDocumentUnits(pipeSegment.Roughness, m_document);
            xPipeSegment.Add(new XAttribute(XName.Get("roughness"), roughnessInDocumentUnits.ToString("r")));

            foreach (MEPSize size in pipeSegment.GetSizes())
                xPipeSegment.Add(CreateXmlFromMEPSize(size, m_document));

            return xPipeSegment;

        }


        /// <summary>
        /// Create Xml from an MEPSize
        /// </summary>
        /// <param name="size"></param>
        /// <param name="document"></param>
        /// <returns></returns>
        private static XElement CreateXmlFromMEPSize(MEPSize size, Autodesk.Revit.DB.Document document)
        {
            XElement xMEPSize = new XElement(XName.Get("MEPSize"));

            xMEPSize.Add(new XAttribute(XName.Get("innerDiameter"), (Utility.Convert.ConvertValueDocumentUnits(size.InnerDiameter, document)).ToString()));
            xMEPSize.Add(new XAttribute(XName.Get("nominalDiameter"), (Utility.Convert.ConvertValueDocumentUnits(size.NominalDiameter, document)).ToString()));
            xMEPSize.Add(new XAttribute(XName.Get("outerDiameter"), (Utility.Convert.ConvertValueDocumentUnits(size.OuterDiameter, document)).ToString()));
            xMEPSize.Add(new XAttribute(XName.Get("usedInSizeLists"), size.UsedInSizeLists));
            xMEPSize.Add(new XAttribute(XName.Get("usedInSizing"), size.UsedInSizing));
            return xMEPSize;
        }
         /// <summary>
        /// Create Xml from a RoutingPreferenceManager
        /// </summary>
        /// <param name="routingPreferenceManager"></param>
        /// <returns></returns>
        private XElement CreateXmlFromRoutingPreferenceManager(RoutingPreferenceManager routingPreferenceManager)
        {
            XElement xRoutingPreferenceManager = new XElement(XName.Get("RoutingPreferenceManager"));

            xRoutingPreferenceManager.Add(new XAttribute(XName.Get("pipeTypeName"), GetPipeTypeNameById(routingPreferenceManager.OwnerId)));

            xRoutingPreferenceManager.Add(new XAttribute(XName.Get("preferredJunctionType"), routingPreferenceManager.PreferredJunctionType.ToString()));

            for (int indexCrosses = 0; indexCrosses != routingPreferenceManager.GetNumberOfRules(RoutingPreferenceRuleGroupType.Crosses); indexCrosses++)
            {
                xRoutingPreferenceManager.Add(createXmlFromRoutingPreferenceRule(routingPreferenceManager.GetRule(RoutingPreferenceRuleGroupType.Crosses, indexCrosses), RoutingPreferenceRuleGroupType.Crosses));
            }

            for (int indexElbows = 0; indexElbows != routingPreferenceManager.GetNumberOfRules(RoutingPreferenceRuleGroupType.Elbows); indexElbows++)
            {
                xRoutingPreferenceManager.Add(createXmlFromRoutingPreferenceRule(routingPreferenceManager.GetRule(RoutingPreferenceRuleGroupType.Elbows, indexElbows), RoutingPreferenceRuleGroupType.Elbows));
            }

            for (int indexSegments = 0; indexSegments != routingPreferenceManager.GetNumberOfRules(RoutingPreferenceRuleGroupType.Segments); indexSegments++)
            {
                xRoutingPreferenceManager.Add(createXmlFromRoutingPreferenceRule(routingPreferenceManager.GetRule(RoutingPreferenceRuleGroupType.Segments, indexSegments), RoutingPreferenceRuleGroupType.Segments));
            }

            for (int indexJunctions = 0; indexJunctions != routingPreferenceManager.GetNumberOfRules(RoutingPreferenceRuleGroupType.Junctions); indexJunctions++)
            {
                xRoutingPreferenceManager.Add(createXmlFromRoutingPreferenceRule(routingPreferenceManager.GetRule(RoutingPreferenceRuleGroupType.Junctions, indexJunctions), RoutingPreferenceRuleGroupType.Junctions));
            }

            for (int indexTransitions = 0; indexTransitions != routingPreferenceManager.GetNumberOfRules(RoutingPreferenceRuleGroupType.Transitions); indexTransitions++)
            {
                xRoutingPreferenceManager.Add(createXmlFromRoutingPreferenceRule(routingPreferenceManager.GetRule(RoutingPreferenceRuleGroupType.Transitions, indexTransitions), RoutingPreferenceRuleGroupType.Transitions));
            }

            for (int indexUnions = 0; indexUnions != routingPreferenceManager.GetNumberOfRules(RoutingPreferenceRuleGroupType.Unions); indexUnions++)
            {
                xRoutingPreferenceManager.Add(createXmlFromRoutingPreferenceRule(routingPreferenceManager.GetRule(RoutingPreferenceRuleGroupType.Unions, indexUnions), RoutingPreferenceRuleGroupType.Unions));
            }

            for (int indexMechanicalJoints = 0; indexMechanicalJoints != routingPreferenceManager.GetNumberOfRules(RoutingPreferenceRuleGroupType.MechanicalJoints); indexMechanicalJoints++)
            {
                xRoutingPreferenceManager.Add(createXmlFromRoutingPreferenceRule(routingPreferenceManager.GetRule(RoutingPreferenceRuleGroupType.MechanicalJoints, indexMechanicalJoints), RoutingPreferenceRuleGroupType.MechanicalJoints));
            }


            return xRoutingPreferenceManager;
        }

        /// <summary>
        /// Create Xml from a RoutingPreferenceRule
        /// </summary>
        /// <param name="rule"></param>
        /// <param name="groupType"></param>
        /// <returns></returns>
        private XElement createXmlFromRoutingPreferenceRule(RoutingPreferenceRule rule, RoutingPreferenceRuleGroupType groupType)
        {
            XElement xRoutingPreferenceRule = new XElement(XName.Get("RoutingPreferenceRule"));
            xRoutingPreferenceRule.Add(new XAttribute(XName.Get("description"), rule.Description));
            xRoutingPreferenceRule.Add(new XAttribute(XName.Get("ruleGroup"), groupType.ToString()));
            if (rule.NumberOfCriteria >= 1)
            {
                PrimarySizeCriterion psc = rule.GetCriterion(0) as PrimarySizeCriterion;

                if (psc.IsEqual(PrimarySizeCriterion.All()))
                {
                    xRoutingPreferenceRule.Add(new XAttribute(XName.Get("minimumSize"), "All"));
                }
                else
                   if (psc.IsEqual(PrimarySizeCriterion.None()))
                {
                    xRoutingPreferenceRule.Add(new XAttribute(XName.Get("minimumSize"), "None"));
                }
                else  //Only specify "maximumSize" if not specifying "All" or "None" for minimum size, just like in the UI.
                {

                    xRoutingPreferenceRule.Add(new XAttribute(XName.Get("minimumSize"), (Utility.Convert.ConvertValueDocumentUnits(psc.MinimumSize, m_document)).ToString()));
                    xRoutingPreferenceRule.Add(new XAttribute(XName.Get("maximumSize"), (Utility.Convert.ConvertValueDocumentUnits(psc.MaximumSize, m_document)).ToString()));
                }
            }
            else
            {
                xRoutingPreferenceRule.Add(new XAttribute(XName.Get("minimumSize"), "All"));
            }

            if (groupType == RoutingPreferenceRuleGroupType.Segments)
            {
                xRoutingPreferenceRule.Add(new XAttribute(XName.Get("partName"), GetSegmentNameById(rule.MEPPartId)));
            }
            else
                xRoutingPreferenceRule.Add(new XAttribute(XName.Get("partName"), GetFittingNameById(rule.MEPPartId)));

            return xRoutingPreferenceRule;
        }
        #endregion



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
