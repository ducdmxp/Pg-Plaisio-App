using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace Convert2DTo3D
{
    public static class ProjectShareParameterUtils
    {
        public static List<BuiltInCategory> BIC_AllowsBoundParametersAsType()
        {
            return new List<BuiltInCategory>(){
                     ///<summary>Analytical Links</summary>
                      BuiltInCategory.OST_LinksAnalytical,
                     ///<summary>Structural Connections</summary>
                      BuiltInCategory.OST_StructConnections,
                     ///<summary>Structural Fabric Areas</summary>
                     BuiltInCategory.OST_FabricAreas,
                     ///<summary>Structural Fabric Reinforcement</summary>
                     BuiltInCategory.OST_FabricReinforcement,
                     ///<summary>Rebar Shape</summary>
                      BuiltInCategory.OST_RebarShape,
                     ///<summary>Structural Path Reinforcement</summary>
                      BuiltInCategory.OST_PathRein,
                     ///<summary>Structural Area Reinforcement</summary>
                     BuiltInCategory.OST_AreaRein,
                     ///<summary>Structural Rebar</summary>
                      BuiltInCategory.OST_Rebar,
                     ///<summary>Pipe Placeholders</summary>
                      BuiltInCategory.OST_PlaceHolderPipes,
                     ///<summary>Duct Placeholders</summary>
                      BuiltInCategory.OST_PlaceHolderDucts,
                     ///<summary>Cable Tray Runs</summary>
                      BuiltInCategory.OST_CableTrayRun,
                     ///<summary>Conduit Runs</summary>
                     BuiltInCategory.OST_ConduitRun,
                     ///<summary>Conduits</summary>
                     BuiltInCategory.OST_Conduit,
                     ///<summary>Cable Trays</summary>
                      BuiltInCategory.OST_CableTray,
                     ///<summary>Conduit Fittings</summary>
                      BuiltInCategory.OST_ConduitFitting,
                     ///<summary>Cable Tray Fittings</summary>
                      BuiltInCategory.OST_CableTrayFitting,
                     ///<summary>Duct Linings</summary>
                      BuiltInCategory.OST_DuctLinings,
                     ///<summary>Duct Insulations</summary>
                     BuiltInCategory.OST_DuctInsulations,
                     ///<summary>Pipe Insulations</summary>
                     BuiltInCategory.OST_PipeInsulations,
                     ///<summary>Switch System</summary>
                     BuiltInCategory.OST_SwitchSystem,
                     ///<summary>Sprinklers</summary>
                     BuiltInCategory.OST_Sprinklers,
                     ///<summary>Lighting Devices</summary>
                      BuiltInCategory.OST_LightingDevices,
                     ///<summary>Fire Alarm Devices</summary>
                      BuiltInCategory.OST_FireAlarmDevices,
                     ///<summary>Data Devices</summary>
                      BuiltInCategory.OST_DataDevices,
                     ///<summary>Communication Devices</summary>
                      BuiltInCategory.OST_CommunicationDevices,
                     ///<summary>Security Devices</summary>
                      BuiltInCategory.OST_SecurityDevices,
                     ///<summary>Nurse Call Devices</summary>
                      BuiltInCategory.OST_NurseCallDevices,
                     ///<summary>Telephone Devices</summary>
                      BuiltInCategory.OST_TelephoneDevices,
                     ///<summary>Pipe Accessories</summary>
                      BuiltInCategory.OST_PipeAccessory,
                     ///<summary>Flex Pipes</summary>
                      BuiltInCategory.OST_FlexPipeCurves,
                     ///<summary>Pipe Fittings</summary>
                      BuiltInCategory.OST_PipeFitting,
                     ///<summary>Pipes</summary>
                      BuiltInCategory.OST_PipeCurves,
                     ///<summary>Piping Systems</summary>
                      BuiltInCategory.OST_PipingSystem,
                     ///<summary>Wires</summary>
                      BuiltInCategory.OST_Wire,
                     ///<summary>Flex Ducts</summary>
                      BuiltInCategory.OST_FlexDuctCurves,
                     ///<summary>Duct Accessories</summary>
                      BuiltInCategory.OST_DuctAccessory,
                     ///<summary>Duct Systems</summary>
                      BuiltInCategory.OST_DuctSystem,
                     ///<summary>Air Terminals</summary>
                      BuiltInCategory.OST_DuctTerminal,
                     ///<summary>Duct Fittings</summary>
                      BuiltInCategory.OST_DuctFitting,
                     ///<summary>Ducts</summary>
                      BuiltInCategory.OST_DuctCurves,
                     ///<summary>Mass</summary>
                      BuiltInCategory.OST_Mass,
                     ///<summary>Detail Items</summary>
                      BuiltInCategory.OST_DetailComponents,
                     ///<summary>Floors.Slab Edges</summary>
                      BuiltInCategory.OST_EdgeSlab,
                     ///<summary>Roofs.Gutters</summary>
                      BuiltInCategory.OST_Gutter,
                     ///<summary>Roofs.Fascias</summary>
                      BuiltInCategory.OST_Fascia,
                     ///<summary>Planting</summary>
                     BuiltInCategory.OST_Planting,
                     ///<summary>Structural Stiffeners</summary>
                      BuiltInCategory.OST_StructuralStiffener,
                     ///<summary>Specialty Equipment</summary>
                      BuiltInCategory.OST_SpecialityEquipment,
                     ///<summary>Topography</summary>
                      BuiltInCategory.OST_Topography,
                     ///<summary>Structural Trusses</summary>
                      BuiltInCategory.OST_StructuralTruss,
                     ///<summary>Structural Columns</summary>
                      BuiltInCategory.OST_StructuralColumns,
                     ///<summary>Structural Beam Systems</summary>
                      BuiltInCategory.OST_StructuralFramingSystem,
                     ///<summary>Structural Framing</summary>
                      BuiltInCategory.OST_StructuralFraming,
                     ///<summary>Structural Foundations</summary>
                      BuiltInCategory.OST_StructuralFoundation,
                     ///<summary>Site.Property Line Segments</summary>
                      BuiltInCategory.OST_SitePropertyLineSegment,
                     ///<summary>Site.Property Lines</summary>
                      BuiltInCategory.OST_SiteProperty,
                     ///<summary>Site.Pads</summary>
                      BuiltInCategory.OST_BuildingPad,
                     ///<summary>Site</summary>
                      BuiltInCategory.OST_Site,
                     ///<summary>Parking</summary>
                      BuiltInCategory.OST_Parking,
                     ///<summary>Plumbing Fixtures</summary>
                      BuiltInCategory.OST_PlumbingFixtures,
                     ///<summary>Mechanical Equipment</summary>
                      BuiltInCategory.OST_MechanicalEquipment,
                     ///<summary>Lighting Fixtures</summary>
                      BuiltInCategory.OST_LightingFixtures,
                     ///<summary>Furniture Systems</summary>
                      BuiltInCategory.OST_FurnitureSystems,
                     ///<summary>Electrical Fixtures</summary>
                      BuiltInCategory.OST_ElectricalFixtures,
                     ///<summary>Electrical Equipment</summary>
                     BuiltInCategory.OST_ElectricalEquipment,
                     ///<summary>Casework</summary>
                      BuiltInCategory.OST_Casework,
                     ///<summary>Railings.Terminations</summary>
                      BuiltInCategory.OST_RailingTermination,

                     /////<summary>Railings.Supports</summary>
                     // BuiltInCategory.OST_RailingSupport,

                     ///<summary>Railings.Handrails</summary>
                      BuiltInCategory.OST_RailingHandRail,
                     ///<summary>Railings.Top Rails</summary>
                     BuiltInCategory.OST_RailingTopRail,
                     ///<summary>Stairs.Landings</summary>
                      BuiltInCategory.OST_StairsLandings,
                     ///<summary>Stairs.Runs</summary>
                      BuiltInCategory.OST_StairsRuns,
                     ///<summary>Curtain Systems</summary>
                      BuiltInCategory.OST_CurtaSystem,
                     ///<summary>Assemblies</summary>
                     BuiltInCategory.OST_Assemblies,
                     ///<summary>Levels</summary>
                      BuiltInCategory.OST_Levels,
                     ///<summary>Grids</summary>
                      BuiltInCategory.OST_Grids,
                     ///<summary>Walls.Wall Sweeps</summary>
                      BuiltInCategory.OST_Cornices,
                     ///<summary>Ramps</summary>
                      BuiltInCategory.OST_Ramps,
                     ///<summary>Curtain Wall Mullions</summary>
                      BuiltInCategory.OST_CurtainWallMullions,
                     ///<summary>Curtain Panels</summary>
                      BuiltInCategory.OST_CurtainWallPanels,
                     ///<summary>Generic Models</summary>
                      BuiltInCategory.OST_GenericModel,
                     ///<summary>Railings</summary>
                      BuiltInCategory.OST_StairsRailing,
                     ///<summary>Stairs</summary>
                      BuiltInCategory.OST_Stairs,
                     ///<summary>Columns</summary>
                      BuiltInCategory.OST_Columns,
                     ///<summary>Furniture</summary>
                      BuiltInCategory.OST_Furniture,
                     ///<summary>Ceilings</summary>
                      BuiltInCategory.OST_Ceilings,
                     ///<summary>Roofs</summary>
                     BuiltInCategory.OST_Roofs,
                     ///<summary>Floors</summary>
                      BuiltInCategory.OST_Floors,
                     ///<summary>Doors</summary>
                      BuiltInCategory.OST_Doors,
                     ///<summary>Windows</summary>
                      BuiltInCategory.OST_Windows,
                     ///<summary>Walls</summary>
                      BuiltInCategory.OST_Walls,
                     ///<summary>Analytical Pipe Connections</summary>
                     BuiltInCategory.OST_AnalyticalPipeConnections,
                     ///<summary>Entourage</summary>
                     BuiltInCategory.OST_Entourage,
                     ///<summary>Multi-segmented Grid</summary>
                     BuiltInCategory.OST_GridChains,
                     ///<summary>MEP Fabrication Containment</summary>
                     BuiltInCategory.OST_FabricationContainment,
                     ///<summary>MEP Fabrication Ductwork</summary>
                     BuiltInCategory.OST_FabricationDuctwork,
                     ///<summary>MEP Fabrication Hangers</summary>
                     BuiltInCategory.OST_FabricationHangers,
                     ///<summary>MEP Fabrication Pipework</summary>
                      BuiltInCategory.OST_FabricationPipework,
                     ///<summary>Mechanical Equipment Sets</summary>
                     BuiltInCategory.OST_MechanicalEquipmentSet,
                     ///<summary>Model Groups</summary>
                     BuiltInCategory.OST_IOSModelGroups,
                     ///<summary>RVT Links</summary>
                     BuiltInCategory.OST_RvtLinks,
                     ///<summary>Roof Soffits</summary>
                     BuiltInCategory.OST_RoofSoffit,

                     /////<summary>Supports</summary>
                     //BuiltInCategory.OST_StairsSupports,

                     //<summary>Fabric Wire</summary>
                     BuiltInCategory.OST_FabricReinforcementWire,
                     ///<summary>Structural Rebar Couplers</summary>
                     BuiltInCategory.OST_Coupler,
                     ///<summary>Topography Links</summary>
                     BuiltInCategory.OST_TopographyLink,
#if  REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024
                ///<summary>Abutments</summary>
                BuiltInCategory.OST_BridgeAbutments,
                ///<summary>Abutment Foundations</summary>
                BuiltInCategory.OST_AbutmentFoundations,
                ///<summary>Abutment Piles</summary>
                BuiltInCategory.OST_AbutmentPiles,
                ///<summary>Abutment Walls</summary>
                BuiltInCategory.OST_AbutmentWalls,
                ///<summary>Approach Slabs</summary>
                BuiltInCategory.OST_ApproachSlabs,
                ///<summary>Alignments</summary>
                BuiltInCategory.OST_Alignments,
                ///<summary>Bearings</summary>
                BuiltInCategory.OST_BridgeBearings,
                ///<summary>Bridge Cables</summary>
                BuiltInCategory.OST_BridgeCables,
                ///<summary>Bridge Decks</summary>
                BuiltInCategory.OST_BridgeDecks,
                ///<summary>Bridge Framing</summary>
                BuiltInCategory.OST_BridgeFraming,
                ///<summary>Arches</summary>
                BuiltInCategory.OST_BridgeArches,
                ///<summary>Cross Bracing</summary>
                BuiltInCategory.OST_BridgeFramingCrossBracing,
                ///<summary>Diaphragms</summary>
                BuiltInCategory.OST_BridgeFramingDiaphragms,
                ///<summary>Girders</summary>
                BuiltInCategory.OST_BridgeGirders,
                ///<summary>Trusses</summary>
                BuiltInCategory.OST_BridgeFramingTrusses,
                ///<summary>Expansion Joints</summary>
                BuiltInCategory.OST_ExpansionJoints,
                ///<summary>Piers</summary>
                BuiltInCategory.OST_BridgePiers,
                ///<summary>Pier Caps</summary>
                BuiltInCategory.OST_PierCaps,
                ///<summary>Pier Columns</summary>
                BuiltInCategory.OST_PierColumns,
                ///<summary>Pier Foundations</summary>
                BuiltInCategory.OST_BridgeFoundations ,
                ///<summary>Pier Piles</summary>
                BuiltInCategory.OST_PierPiles ,
                ///<summary>Pier Towers</summary>
                BuiltInCategory.OST_BridgeTowers,
                ///<summary>Pier Walls</summary>
                BuiltInCategory.OST_PierWalls,
                ///<summary>Roads</summary>
                BuiltInCategory.OST_Roads,
                ///<summary>Structural Tendons</summary>
                BuiltInCategory.OST_StructuralTendons,
                ///<summary>Vibration Management</summary>
                BuiltInCategory.OST_VibrationManagement,
                ///<summary>Vibration Dampers</summary>
                BuiltInCategory.OST_VibrationDampers,
                ///<summary>Vibration Isolators</summary>
                BuiltInCategory.OST_VibrationIsolators,

#endif
#if REVIT2022 || REVIT2023 || REVIT2024
                ///<summary> Fire Protection</summary>
                BuiltInCategory.OST_FireProtection,
                ///<summary>Food Service Equipment</summary>
                BuiltInCategory.OST_FoodServiceEquipment,
                ///<summary>Hardscape</summary>
                BuiltInCategory.OST_Hardscape,
                ///<summary>Medical Equipment</summary>
                BuiltInCategory.OST_MedicalEquipment,
                ///<summary>Signage</summary>
                BuiltInCategory.OST_Signage,
                ///<summary>Temporary Structures</summary>
                BuiltInCategory.OST_TemporaryStructure ,
                ///<summary>Vertical Circulation</summary>
                BuiltInCategory.OST_VerticalCirculation,
                ///<summary>Audio Visual Devices</summary>
                BuiltInCategory.OST_AudioVisualDevices,
#endif
#if REVIT2023 || REVIT2024
                 ///<summary>Mechanical Control Devices</summary>
                 BuiltInCategory.OST_MechanicalControlDevices,
                 ///<summary>MEP Ancillary Framing</summary>
                 BuiltInCategory.OST_MEPAncillaryFraming,
                 ///<summary>MEP Fabrication Ductwork Stiffeners</summary>
                 BuiltInCategory.OST_FabricationDuctworkStiffeners,
                 ///<summary>Plumbing Equipment</summary>
                 BuiltInCategory.OST_PlumbingEquipment,
                 ///<summary>Toposolid</summary>
                 BuiltInCategory.OST_Toposolid,
                ///<summary>Toposolid Links</summary>
                 BuiltInCategory.OST_ToposolidLink,
#endif
#if REVIT2024
                ///<summary>Data Exchanges</summary>
                 BuiltInCategory.OST_DataExchanges,
#endif
            };
        }

        public static List<BuiltInCategory> BIC_AllowsBoundParametersAsInstance()
        {
            return new List<BuiltInCategory>()
            {
                ///<summary>Analytical Links</summary>
                BuiltInCategory.OST_LinksAnalytical,
                ///<summary>Analytical Nodes</summary>
                BuiltInCategory.OST_AnalyticalNodes,
                ///<summary>Analytical Foundation Slabs</summary>
                BuiltInCategory.OST_FoundationSlabAnalytical,
                ///<summary>Analytical Wall Foundations</summary>
                 BuiltInCategory.OST_WallFoundationAnalytical,
                ///<summary>Analytical Isolated Foundations</summary>
                 BuiltInCategory.OST_IsolatedFoundationAnalytical,
                ///<summary>Analytical Walls</summary>
                BuiltInCategory.OST_WallAnalytical,
                ///<summary>Analytical Floors</summary>
                BuiltInCategory.OST_FloorAnalytical,
                ///<summary>Analytical Columns</summary>
                BuiltInCategory.OST_ColumnAnalytical,
                ///<summary>Analytical Braces</summary>
                 BuiltInCategory.OST_BraceAnalytical,
                ///<summary>Analytical Beams</summary>
                 BuiltInCategory.OST_BeamAnalytical,
                ///<summary>Structural Connections</summary>
                 BuiltInCategory.OST_StructConnections,
                ///<summary>Structural Fabric Areas</summary>
                 BuiltInCategory.OST_FabricAreas,
                ///<summary>Structural Fabric Reinforcement</summary>
                BuiltInCategory.OST_FabricReinforcement,
                ///<summary>Rebar Shape</summary>
                 BuiltInCategory.OST_RebarShape,
                ///<summary>Structural Path Reinforcement</summary>
                 BuiltInCategory.OST_PathRein,
                ///<summary>Structural Area Reinforcement</summary>
                BuiltInCategory.OST_AreaRein,
                ///<summary>Structural Rebar</summary>
                BuiltInCategory.OST_Rebar,
                ///<summary>Analytical Spaces</summary>
                BuiltInCategory.OST_AnalyticSpaces,
                ///<summary>Pipe Placeholders</summary>
                BuiltInCategory.OST_PlaceHolderPipes,
                ///<summary>Duct Placeholders</summary>
                BuiltInCategory.OST_PlaceHolderDucts,
                ///<summary>Cable Tray Runs</summary>
                 BuiltInCategory.OST_CableTrayRun,
                ///<summary>Conduit Runs</summary>
                BuiltInCategory.OST_ConduitRun,
                ///<summary>Conduits</summary>
                BuiltInCategory.OST_Conduit,
                ///<summary>Cable Trays</summary>
                 BuiltInCategory.OST_CableTray,
                ///<summary>Conduit Fittings</summary>
                BuiltInCategory.OST_ConduitFitting,
                ///<summary>Cable Tray Fittings</summary>
                 BuiltInCategory.OST_CableTrayFitting,
                ///<summary>Duct Linings</summary>
                 BuiltInCategory.OST_DuctLinings,
                ///<summary>Duct Insulations</summary>
                 BuiltInCategory.OST_DuctInsulations,
                ///<summary>Pipe Insulations</summary>
                BuiltInCategory.OST_PipeInsulations,
                ///<summary>HVAC Zones</summary>
                 BuiltInCategory.OST_HVAC_Zones,
                ///<summary>Switch System</summary>
                 BuiltInCategory.OST_SwitchSystem,
                ///<summary>Sprinklers</summary>
                BuiltInCategory.OST_Sprinklers,
                ///<summary>Analytical Surfaces</summary>
                 BuiltInCategory.OST_AnalyticSurfaces,
                ///<summary>Lighting Devices</summary>
               BuiltInCategory.OST_LightingDevices,
                ///<summary>Fire Alarm Devices</summary>
                 BuiltInCategory.OST_FireAlarmDevices,
                ///<summary>Data Devices</summary>
                BuiltInCategory.OST_DataDevices,
                ///<summary>Communication Devices</summary>
                 BuiltInCategory.OST_CommunicationDevices,
                ///<summary>Security Devices</summary>
                 BuiltInCategory.OST_SecurityDevices,
                ///<summary>Nurse Call Devices</summary>
                 BuiltInCategory.OST_NurseCallDevices,
                ///<summary>Telephone Devices</summary>
                 BuiltInCategory.OST_TelephoneDevices,
                ///<summary>Pipe Accessories</summary>
                BuiltInCategory.OST_PipeAccessory,
                ///<summary>Flex Pipes</summary>
                 BuiltInCategory.OST_FlexPipeCurves,
                ///<summary>Pipe Fittings</summary>
                BuiltInCategory.OST_PipeFitting,
                ///<summary>Pipes</summary>
                 BuiltInCategory.OST_PipeCurves,
                ///<summary>Piping Systems</summary>
                 BuiltInCategory.OST_PipingSystem,
                ///<summary>Wires</summary>
                 BuiltInCategory.OST_Wire,
                ///<summary>Electrical Circuits</summary>
                 BuiltInCategory.OST_ElectricalCircuit,
                ///<summary>Flex Ducts</summary>
                 BuiltInCategory.OST_FlexDuctCurves,
                ///<summary>Duct Accessories</summary>
                 BuiltInCategory.OST_DuctAccessory,
                ///<summary>Duct Systems</summary>
                 BuiltInCategory.OST_DuctSystem,
                ///<summary>Air Terminals</summary>
                 BuiltInCategory.OST_DuctTerminal,
                ///<summary>Duct Fittings</summary>
                 BuiltInCategory.OST_DuctFitting,
                ///<summary>Ducts</summary>
                 BuiltInCategory.OST_DuctCurves,
                ///<summary>Structural Internal Loads.Internal Area Loads</summary>
                 BuiltInCategory.OST_InternalAreaLoads,
                ///<summary>Structural Internal Loads.Internal Line Loads</summary>
                 BuiltInCategory.OST_InternalLineLoads,
                ///<summary>Structural Internal Loads.Internal Point Loads</summary>
                 BuiltInCategory.OST_InternalPointLoads,
                ///<summary>Structural Loads.Area Loads</summary>
                 BuiltInCategory.OST_AreaLoads,
                ///<summary>Structural Loads.Line Loads</summary>
                 BuiltInCategory.OST_LineLoads,
                ///<summary>Structural Loads.Point Loads</summary>
                 BuiltInCategory.OST_PointLoads,
                ///<summary>Spaces</summary>
                 BuiltInCategory.OST_MEPSpaces,
                ///<summary>Mass.Mass Opening</summary>
                 BuiltInCategory.OST_MassOpening,
                ///<summary>Mass.Mass Skylight</summary>
                 BuiltInCategory.OST_MassSkylights,
                ///<summary>Mass.Mass Roof</summary>
                 BuiltInCategory.OST_MassRoof,
                ///<summary>Mass.Mass Exterior Wall</summary>
                 BuiltInCategory.OST_MassExteriorWall,
                ///<summary>Mass.Mass Interior Wall</summary>
                 BuiltInCategory.OST_MassInteriorWall,
                ///<summary>Mass.Mass Zone</summary>
                 BuiltInCategory.OST_MassZone,
                ///<summary>Mass.Mass Floor</summary>
                 BuiltInCategory.OST_MassFloor,
                ///<summary>Mass</summary>
                BuiltInCategory.OST_Mass,
                ///<summary>Areas</summary>
                 BuiltInCategory.OST_Areas,
                ///<summary>Project Information</summary>
                 BuiltInCategory.OST_ProjectInformation,
                ///<summary>Sheets</summary>
                 BuiltInCategory.OST_Sheets,
                ///<summary>Detail Items</summary>
                 BuiltInCategory.OST_DetailComponents,
                ///<summary>Floors.Slab Edges</summary>
                 BuiltInCategory.OST_EdgeSlab,
                ///<summary>Roofs.Gutters</summary>
                 BuiltInCategory.OST_Gutter,
                ///<summary>Roofs.Fascias</summary>
                 BuiltInCategory.OST_Fascia,
                ///<summary>Planting</summary>
                 BuiltInCategory.OST_Planting,
                ///<summary>Structural Stiffeners</summary>
                 BuiltInCategory.OST_StructuralStiffener,
                ///<summary>Specialty Equipment</summary>
                 BuiltInCategory.OST_SpecialityEquipment,
                ///<summary>Topography</summary>
                 BuiltInCategory.OST_Topography,
                ///<summary>Structural Trusses</summary>
                 BuiltInCategory.OST_StructuralTruss,
                ///<summary>Structural Columns</summary>
                 BuiltInCategory.OST_StructuralColumns,
                ///<summary>Structural Beam Systems</summary>
                 BuiltInCategory.OST_StructuralFramingSystem,
                ///<summary>Structural Framing</summary>
                 BuiltInCategory.OST_StructuralFraming,
                ///<summary>Structural Foundations</summary>
                 BuiltInCategory.OST_StructuralFoundation,
                ///<summary>Site.Property Line Segments</summary>
                 BuiltInCategory.OST_SitePropertyLineSegment,
                ///<summary>Site.Property Lines</summary>
                 BuiltInCategory.OST_SiteProperty,
                ///<summary>Site.Pads</summary>
                 BuiltInCategory.OST_BuildingPad,
                ///<summary>Site</summary>
                 BuiltInCategory.OST_Site,
                ///<summary>Roads</summary>
                 BuiltInCategory.OST_Roads,
                ///<summary>Parking</summary>
                 BuiltInCategory.OST_Parking,
                ///<summary>Plumbing Fixtures</summary>
                 BuiltInCategory.OST_PlumbingFixtures,
                ///<summary>Mechanical Equipment</summary>
                 BuiltInCategory.OST_MechanicalEquipment,
                ///<summary>Lighting Fixtures</summary>
                 BuiltInCategory.OST_LightingFixtures,
                ///<summary>Furniture Systems</summary>
                 BuiltInCategory.OST_FurnitureSystems,
                ///<summary>Electrical Fixtures</summary>
                 BuiltInCategory.OST_ElectricalFixtures,
                ///<summary>Electrical Equipment</summary>
                 BuiltInCategory.OST_ElectricalEquipment,
                ///<summary>Casework</summary>
                 BuiltInCategory.OST_Casework,
                ///<summary>Shaft Openings</summary>
                 BuiltInCategory.OST_ShaftOpening,
                ///<summary>Railings.Terminations</summary>
                 BuiltInCategory.OST_RailingTermination,

                /////<summary>Railings.Supports</summary>
                // BuiltInCategory.OST_RailingSupport,

                ///<summary>Railings.Handrails</summary>
                 BuiltInCategory.OST_RailingHandRail,
                ///<summary>Railings.Top Rails</summary>
                 BuiltInCategory.OST_RailingTopRail,
                ///<summary>Stairs.Landings</summary>
                 BuiltInCategory.OST_StairsLandings,
                ///<summary>Stairs.Runs</summary>
                 BuiltInCategory.OST_StairsRuns,
                ///<summary>Materials</summary>
                 BuiltInCategory.OST_Materials,
                ///<summary>Curtain Systems</summary>
                 BuiltInCategory.OST_CurtaSystem,
                ///<summary>Views</summary>
                 BuiltInCategory.OST_Views,
                ///<summary>Parts</summary>
                 BuiltInCategory.OST_Parts,
                ///<summary>Assemblies</summary>
                 BuiltInCategory.OST_Assemblies,
                ///<summary>Levels</summary>
                 BuiltInCategory.OST_Levels,
                ///<summary>Grids</summary>
                 BuiltInCategory.OST_Grids,
                ///<summary>Walls.Wall Sweeps</summary>
                 BuiltInCategory.OST_Cornices,
                ///<summary>Ramps</summary>
                 BuiltInCategory.OST_Ramps,
                ///<summary>Curtain Wall Mullions</summary>
                 BuiltInCategory.OST_CurtainWallMullions,
                ///<summary>Curtain Panels</summary>
                 BuiltInCategory.OST_CurtainWallPanels,
                ///<summary>Rooms</summary>
                 BuiltInCategory.OST_Rooms,
                ///<summary>Generic Models</summary>
                 BuiltInCategory.OST_GenericModel,
                ///<summary>Railings</summary>
                 BuiltInCategory.OST_StairsRailing,
                ///<summary>Stairs</summary>
                 BuiltInCategory.OST_Stairs,
                ///<summary>Columns</summary>
                 BuiltInCategory.OST_Columns,
                ///<summary>Furniture</summary>
                BuiltInCategory.OST_Furniture,
                ///<summary>Ceilings</summary>
                 BuiltInCategory.OST_Ceilings,
                ///<summary>Roofs</summary>
                 BuiltInCategory.OST_Roofs,
                ///<summary>Floors</summary>
                BuiltInCategory.OST_Floors,
                ///<summary>Doors</summary>
                 BuiltInCategory.OST_Doors,
                ///<summary>Windows</summary>
                 BuiltInCategory.OST_Windows,
                ///<summary>Walls</summary>
                 BuiltInCategory.OST_Walls,
                 ///<summary>Entourage</summary>
                 BuiltInCategory.OST_Entourage,
                 ///<summary>Path of Travel Lines</summary>
                 BuiltInCategory.OST_PathOfTravelLines,
                 ///<summary>Mass Glazing</summary>
                 BuiltInCategory.OST_MassGlazing,
                  ///<summary>Roof Soffit</summary>
                 BuiltInCategory.OST_RoofSoffit,
                 ///<summary>Rvt Links</summary>
                 BuiltInCategory.OST_RvtLinks,
                 ///<summary>Schedules</summary>
                 BuiltInCategory.OST_Schedules,

                 /////<summary>Stairs[Supports]</summary>
                 //BuiltInCategory.OST_StairsSupports,

                 ///<summary>Structural Connections[Anchors]</summary>
                 BuiltInCategory.OST_StructConnectionAnchors,
                 ///<summary>Structural Connections[Bolts]</summary>
                 BuiltInCategory.OST_StructConnectionBolts,
                  ///<summary>Structural Connections[Holes]</summary>
                 BuiltInCategory.OST_StructConnectionHoles,
                 ///<summary>Structural Connections[Modifiers]</summary>
                 BuiltInCategory.OST_StructConnectionModifiers,
                 ///<summary>Structural Connections[Others]</summary>
                 BuiltInCategory.OST_StructConnectionOthers,
                 ///<summary>Structural Connections[Plates]</summary>
                 BuiltInCategory.OST_StructConnectionPlates,
                 ///<summary>Structural Connections[Profiles]</summary>
                 BuiltInCategory.OST_StructConnectionProfiles,
                 ///<summary>Structural Connections[ShearStuds]</summary>
                 BuiltInCategory.OST_StructConnectionShearStuds,
                 ///<summary>Structural Connections[Welds]</summary>
                 BuiltInCategory.OST_StructConnectionWelds,
                 ///<summary>Structural Fabric Reinforcement[Fabric Wire]</summary>
                 BuiltInCategory.OST_FabricReinforcementWire,
                  ///<summary>Structural Rebar Coupler</summary>
                 BuiltInCategory.OST_Coupler,
                 ///<summary>Topography Link</summary>
                 BuiltInCategory.OST_TopographyLink,
                  ///<summary>Analytical Pipe Connections</summary>
                 BuiltInCategory.OST_AnalyticalPipeConnections,
                 ///<summary>MEP Fabriccation Containment</summary>
                 BuiltInCategory.OST_FabricationContainment,
                  ///<summary>MEP Fabriccation Ductwork</summary>
                 BuiltInCategory.OST_FabricationDuctwork,
                 ///<summary>MEP Fabriccation Ductwork[Insulation]</summary>
                 BuiltInCategory.OST_FabricationDuctworkInsulation,
                 ///<summary>MEP Fabriccation Ductwork[Lining]</summary>
                 BuiltInCategory.OST_FabricationDuctworkLining,
                 ///<summary>MEP Fabriccation Hangers</summary>
                 BuiltInCategory.OST_FabricationHangers,
                 ///<summary>MEP Fabriccation Pipework</summary>
                 BuiltInCategory.OST_FabricationPipework,
                 ///<summary>MEP Fabriccation Pipework[Insulation]</summary>
                 BuiltInCategory.OST_FabricationPipeworkInsulation,
                 ///<summary>Mechanical Equipment Set</summary>
                 BuiltInCategory.OST_MechanicalEquipmentSet,
                 ///<summary>Zone Equipment</summary>
                 BuiltInCategory.OST_ZoneEquipment,
                 ///<summary>Air Systems</summary>
                 BuiltInCategory.OST_MEPAnalyticalAirLoop,
                   ///<summary>System-Zones</summary>
                 BuiltInCategory.OST_MEPSystemZone,
                  ///<summary>Model group</summary>
                 BuiltInCategory.OST_IOSModelGroups,
                 ///<summary>Water Loops</summary>
                 BuiltInCategory.OST_MEPAnalyticalWaterLoop,
                 ///<summary>Multi-segmented Grid</summary>
                 BuiltInCategory.OST_GridChains,
                 ///<summary>Insulation</summary>
                 BuiltInCategory.OST_FabricationPipeworkInsulation,
#if REVIT2021 || REVIT2022 || REVIT2023 || REVIT2024
                ///<summary>Abutments</summary>
                BuiltInCategory.OST_BridgeAbutments,
                ///<summary>Abutment Foundations</summary>
                BuiltInCategory.OST_AbutmentFoundations,
                ///<summary>Abutment Piles</summary>
                BuiltInCategory.OST_AbutmentPiles,
                ///<summary>Abutment Walls</summary>
                BuiltInCategory.OST_AbutmentWalls,
                ///<summary>Approach Slabs</summary>
                BuiltInCategory.OST_ApproachSlabs,
                ///<summary>Alignments</summary>
                BuiltInCategory.OST_Alignments,
                ///<summary>Bearings</summary>
                BuiltInCategory.OST_BridgeBearings,
                ///<summary>Bridge Cables</summary>
                BuiltInCategory.OST_BridgeCables,
                ///<summary>Bridge Decks</summary>
                BuiltInCategory.OST_BridgeDecks,
                ///<summary>Bridge Framing</summary>
                BuiltInCategory.OST_BridgeFraming,
                ///<summary>Arches</summary>
                BuiltInCategory.OST_BridgeArches,
                ///<summary>Cross Bracing</summary>
                BuiltInCategory.OST_BridgeFramingCrossBracing,
                ///<summary>Diaphragms</summary>
                BuiltInCategory.OST_BridgeFramingDiaphragms,
                ///<summary>Girders</summary>
                BuiltInCategory.OST_BridgeGirders,
                ///<summary>Trusses</summary>
                BuiltInCategory.OST_BridgeFramingTrusses,
                ///<summary>Expansion Joints</summary>
                BuiltInCategory.OST_ExpansionJoints,
                ///<summary>Piers</summary>
                BuiltInCategory.OST_BridgePiers,
                ///<summary>Pier Caps</summary>
                BuiltInCategory.OST_PierCaps,
                ///<summary>Pier Columns</summary>
                BuiltInCategory.OST_PierColumns,
                ///<summary>Pier Foundations</summary>
                BuiltInCategory.OST_BridgeFoundations ,
                ///<summary>Pier Piles</summary>
                BuiltInCategory.OST_PierPiles ,
                ///<summary>Pier Towers</summary>
                BuiltInCategory.OST_BridgeTowers,
                ///<summary>Pier Walls</summary>
                BuiltInCategory.OST_PierWalls,
                ///<summary>Structural Tendons</summary>
                BuiltInCategory.OST_StructuralTendons,
                ///<summary>Vibration Management</summary>
                BuiltInCategory.OST_VibrationManagement ,
                ///<summary>Vibration Dampers</summary>
                BuiltInCategory.OST_VibrationDampers,
                ///<summary>Vibration Isolators</summary>
                BuiltInCategory.OST_VibrationIsolators,
#endif
#if REVIT2022 || REVIT2023 || REVIT2024
                 ///<summary>Audio Visual Devices</summary>
                BuiltInCategory.OST_AudioVisualDevices,
                ///<summary>Fire Protection</summary>
                BuiltInCategory.OST_FireProtection,
                ///<summary>Food Service Equipment</summary>
                BuiltInCategory.OST_FoodServiceEquipment,
                ///<summary>Medical Equipment</summary>
                BuiltInCategory.OST_MedicalEquipment,
                ///<summary>Signage</summary>
                BuiltInCategory.OST_Signage,
                ///<summary>Temporary Structure</summary>
                BuiltInCategory.OST_TemporaryStructure,
                ///<summary>Vertical Circulation</summary>
                BuiltInCategory.OST_VerticalCirculation,
                ///<summary>Hardscape</summary>
                BuiltInCategory.OST_Hardscape,
#endif
#if REVIT2023 || REVIT2024
                ///<summary>Analytical Members</summary>
                BuiltInCategory.OST_AnalyticalMember ,
                ///<summary>Analytical Openings</summary>
                 BuiltInCategory.OST_AnalyticalOpening,
                ///<summary>Analytical Panels</summary>
                 BuiltInCategory.OST_AnalyticalPanel ,
                ///<summary>Electrical Analytical Bus</summary>
                 BuiltInCategory.OST_MEPAnalyticalBus,
                ///<summary>Electrical Analytical Load Set</summary>
                 BuiltInCategory.OST_ElectricalLoadSet,
                ///<summary>Electrical Analytical Loads</summary>
                 BuiltInCategory.OST_ElectricalLoadZoneInstance,
                //<summary>Electrical Analytical Power Source</summary>
                 BuiltInCategory.OST_ElectricalPowerSource,
                ///<summary>Electrical Analytical Transfer Switch</summary>
                 BuiltInCategory.OST_MEPAnalyticalTransferSwitch,
                ///<summary>Electrical Analytical Transformer</summary>
                 BuiltInCategory.OST_ElectricalAnalyticalTransformer,
                ///<summary>Electrical Load Areas</summary>
                 BuiltInCategory.OST_MEPLoadAreas,
                ///<summary>Mechanical Control Devices</summary>
                 BuiltInCategory.OST_MechanicalControlDevices,
                ///<summary>MEP Ancillary Framing</summary>
                 BuiltInCategory.OST_MEPAncillaryFraming,
                ///<summary>MEP Fabrication Ductwork Stiffeners</summary>
                 BuiltInCategory.OST_FabricationDuctworkStiffeners,
                ///<summary>Plumbing Equipment</summary>
                 BuiltInCategory.OST_PlumbingEquipment,
#if !REVIT2023
                ///<summary>Revision Clouds</summary>
                BuiltInCategory.OST_RevisionClouds,
#endif

                ///<summary>Toposolid</summary>
                 BuiltInCategory.OST_Toposolid,
                ///<summary>Toposolid Links</summary>
                 BuiltInCategory.OST_ToposolidLink ,
#endif
#if REVIT2024
                ///<summary>Data Exchanges</summary>
                 BuiltInCategory.OST_DataExchanges,
#endif
            };
        }

#if REVIT2019 || REVIT2020 || REVIT2021
        public static List<BuiltInParameterGroup> GetAllParameterGroups()
        {
            return new List<BuiltInParameterGroup>() {
                ///<summary> Analysis Results</summary>
                BuiltInParameterGroup.PG_ANALYSIS_RESULTS,
                ///<summary> Analytical Alignment</summary>
                BuiltInParameterGroup.PG_ANALYTICAL_ALIGNMENT,
                ///<summary> Analytical Model</summary>
                BuiltInParameterGroup.PG_ANALYTICAL_MODEL,
                ///<summary> Constraints</summary>
                BuiltInParameterGroup.PG_CONSTRAINTS,
                ///<summary> Construction</summary>
                BuiltInParameterGroup.PG_CONSTRUCTION,
                ///<summary>Data </summary>
                BuiltInParameterGroup.PG_DATA,
                ///<summary> Dimensions</summary>
                BuiltInParameterGroup.PG_GEOMETRY,
                ///<summary>Division Geometry </summary>
                BuiltInParameterGroup.PG_DIVISION_GEOMETRY,
                ///<summary>Electrical </summary>
                BuiltInParameterGroup.PG_AELECTRICAL,
                ///<summary>Electrical - Circuiting </summary>
                BuiltInParameterGroup.PG_ELECTRICAL_CIRCUITING,
                ///<summary> Electrical - Lighting</summary>
                BuiltInParameterGroup.PG_ELECTRICAL_LIGHTING,
                ///<summary> Electrical - Loads</summary>
                BuiltInParameterGroup.PG_ELECTRICAL_LOADS,
                ///<summary>Electrical Engineering</summary>
                BuiltInParameterGroup.PG_ELECTRICAL,
                ///<summary>Energy Analysis</summary>
                BuiltInParameterGroup.PG_ENERGY_ANALYSIS,
                ///<summary>Fire Protection</summary>
                BuiltInParameterGroup.PG_FIRE_PROTECTION,
                ///<summary>Forces</summary>
                BuiltInParameterGroup.PG_FORCES,
                ///<summary>General</summary>
                BuiltInParameterGroup.PG_GENERAL,
                ///<summary>Graphics</summary>
                BuiltInParameterGroup.PG_GRAPHICS,
                ///<summary>Green Building Properties</summary>
                BuiltInParameterGroup.PG_GREEN_BUILDING,
                ///<summary>Identity Data</summary>
                BuiltInParameterGroup.PG_IDENTITY_DATA,
                ///<summary>IFC Parameters</summary>
                BuiltInParameterGroup.PG_IFC,
                ///<summary>Layers</summary>
                BuiltInParameterGroup.PG_REBAR_SYSTEM_LAYERS,
                ///<summary>Materials and Finishes</summary>
                BuiltInParameterGroup.PG_MATERIALS,
                ///<summary>Mechanical</summary>
                BuiltInParameterGroup.PG_MECHANICAL,
                ///<summary>Mechanical - Flow</summary>
                BuiltInParameterGroup.PG_MECHANICAL_AIRFLOW,
                ///<summary>Mechanical - Loads</summary>
                BuiltInParameterGroup.PG_MECHANICAL_LOADS,
                ///<summary>Model Properties</summary>
                BuiltInParameterGroup.PG_ADSK_MODEL_PROPERTIES,
                ///<summary>Moments</summary>
                BuiltInParameterGroup.PG_MOMENTS,
                ///<summary>Other</summary>
                BuiltInParameterGroup.INVALID,
                ///<summary>Overall Legend</summary>
                BuiltInParameterGroup.PG_OVERALL_LEGEND,
                ///<summary>Phasing</summary>
                BuiltInParameterGroup.PG_PHASING,
                ///<summary>Photometrics</summary>
                 BuiltInParameterGroup.PG_LIGHT_PHOTOMETRICS,
                ///<summary>Plumbing</summary>
                 BuiltInParameterGroup.PG_PLUMBING,
                ///<summary>Primary End</summary>
                 BuiltInParameterGroup.PG_PRIMARY_END,
                ///<summary>Rebar Set</summary>
                 BuiltInParameterGroup.PG_REBAR_ARRAY,
                ///<summary>Releases / Member Forces</summary>
                 BuiltInParameterGroup.PG_RELEASES_MEMBER_FORCES,
                ///<summary>Secondary End</summary>
                 BuiltInParameterGroup.PG_SECONDARY_END,
                ///<summary>Segments and Fittings</summary>
                 BuiltInParameterGroup.PG_SEGMENTS_FITTINGS,
                ///<summary>Set</summary>
                 BuiltInParameterGroup.PG_COUPLER_ARRAY,
                ///<summary>Slab Shape Edit</summary>
                 BuiltInParameterGroup.PG_SLAB_SHAPE_EDIT,
                ///<summary>Structural</summary>
                 BuiltInParameterGroup.PG_STRUCTURAL,
                ///<summary>Structural Analysis</summary>
                 BuiltInParameterGroup.PG_STRUCTURAL_ANALYSIS,
                ///<summary>Text</summary>
                 BuiltInParameterGroup.PG_TEXT,
                ///<summary>Title Text</summary>
                 BuiltInParameterGroup.PG_TITLE,
                ///<summary>Visibility</summary>
                BuiltInParameterGroup.PG_VISIBILITY
            };
        }
#else

        public static List<ForgeTypeId> GetAllParameterGroups()
        {
            return new List<ForgeTypeId>() {
                ///<summary>Analysis Results</summary>
                GroupTypeId.AnalysisResults,
                ///<summary>Analytical Alignment</summary>
                GroupTypeId.AnalyticalAlignment,
                ///<summary>Analytical Model</summary>
                GroupTypeId.AnalyticalModel,
                ///<summary>Constraints</summary>
                GroupTypeId.Constraints,
                ///<summary>Construction</summary>
                GroupTypeId.Construction,
                ///<summary>Data</summary>
                GroupTypeId.Data,
                ///<summary>Dimensions</summary>
                GroupTypeId.Geometry,
                ///<summary>Division Geometry</summary>
                GroupTypeId.DivisionGeometry,
                ///<summary>Electrical</summary>
                GroupTypeId.Electrical,
                ///<summary>Electrical - Circuiting</summary>
                GroupTypeId.ElectricalCircuiting,
                ///<summary>Electrical - Lighting</summary>
                GroupTypeId.ElectricalLighting,
                ///<summary>Electrical - Loads</summary>
                GroupTypeId.ElectricalLoads,
                ///<summary>ElectricalAnalysis</summary>
                GroupTypeId.ElectricalAnalysis,
                ///<summary>Electrical Engineering</summary>
                GroupTypeId.ElectricalEngineering,
                ///<summary>Energy Analysis</summary>
                GroupTypeId.EnergyAnalysis,
                ///<summary>Fire Protection</summary>
                GroupTypeId.FireProtection,
                ///<summary>Forces</summary>
                GroupTypeId.Forces,
                ///<summary>General</summary>
                GroupTypeId.General,
                ///<summary>Graphics</summary>
                GroupTypeId.Graphics,
                ///<summary>Green Building Properties</summary>
                GroupTypeId.GreenBuilding,
                ///<summary>Identity Data</summary>
                GroupTypeId.IdentityData,
                ///<summary>IFC Parameters</summary>
                GroupTypeId.Ifc,
                ///<summary>Layers</summary>
                GroupTypeId.RebarSystemLayers,
                ///<summary>Materials and Finishes</summary>
                GroupTypeId.Materials,
                ///<summary>Mechanical</summary>
                GroupTypeId.Mechanical,
                ///<summary>Mechanical - Flow</summary>
                GroupTypeId.MechanicalAirflow,
                ///<summary>Mechanical - Loads</summary>
                GroupTypeId.MechanicalLoads,
                ///<summary>Model Properties</summary>
                GroupTypeId.AdskModelProperties,
                ///<summary>Moments</summary>
                GroupTypeId.Moments,
                ///<summary>Overall Legend</summary>
                GroupTypeId.OverallLegend,
                ///<summary>Phasing</summary>
                GroupTypeId.Phasing,
                ///<summary>Photometrics</summary>
                GroupTypeId.LightPhotometrics,
                ///<summary>Plumbing</summary>
                GroupTypeId.Plumbing,
                ///<summary>Primary End</summary>
                GroupTypeId.PrimaryEnd,
                ///<summary>Rebar Set</summary>
                GroupTypeId.RebarArray,
                ///<summary>Releases / Member Forces</summary>
                GroupTypeId.ReleasesMemberForces,
                ///<summary>Secondary End</summary>
                GroupTypeId.SecondaryEnd,
                ///<summary>Segments and Fittings</summary>
                GroupTypeId.SegmentsFittings,
                ///<summary>Set</summary>
                GroupTypeId.CouplerArray,
                ///<summary>Slab Shape Edit</summary>
                GroupTypeId.SlabShapeEdit,
                ///<summary>Structural</summary>
                GroupTypeId.Structural,
                ///<summary>Structural Analysis</summary>
                GroupTypeId.StructuralAnalysis,
                ///<summary>Text</summary>
                GroupTypeId.Text,
                ///<summary>Title Text</summary>
                GroupTypeId.Title,
                ///<summary>Visibility</summary>
                GroupTypeId.Visibility,
#if REVIT2023 || REVIT2024
                ///<summary>Life Safety</summary>
                GroupTypeId.LifeSafety,
#endif
#if REVIT2024
                ///<summary>Structural Section Dimensions</summary>
                GroupTypeId.StructuralSectionDimensions,
                ///<summary>Sub-Division</summary>
                GroupTypeId.ToposolidSubdivision
#endif
            };
        }

#endif
    }
}