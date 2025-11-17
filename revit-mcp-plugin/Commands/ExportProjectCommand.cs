using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace revit_mcp_plugin.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ExportProjectCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;

                // Show save file dialog
                System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog
                {
                    Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                    FilterIndex = 1,
                    RestoreDirectory = true,
                    FileName = $"{doc.Title}_Export_{DateTime.Now:yyyyMMdd_HHmmss}.json",
                    Title = "Export Project to JSON"
                };

                if (saveFileDialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                {
                    return Result.Cancelled;
                }

                string filePath = saveFileDialog.FileName;

                // Collect all elements from the project
                var collector = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType();

                var categories = new List<BuiltInCategory>
                {
                    BuiltInCategory.OST_Walls,
                    BuiltInCategory.OST_Doors,
                    BuiltInCategory.OST_Windows,
                    BuiltInCategory.OST_Furniture,
                    BuiltInCategory.OST_Columns,
                    BuiltInCategory.OST_Floors,
                    BuiltInCategory.OST_Roofs,
                    BuiltInCategory.OST_Stairs,
                    BuiltInCategory.OST_StructuralFraming,
                    BuiltInCategory.OST_Ceilings,
                    BuiltInCategory.OST_Rooms,
                    BuiltInCategory.OST_MEPSpaces,
                    BuiltInCategory.OST_Railings,
                    BuiltInCategory.OST_StructuralColumns,
                    BuiltInCategory.OST_StructuralFoundation,
                    BuiltInCategory.OST_CurtainWallPanels,
                    BuiltInCategory.OST_CurtainWallMullions,
                    BuiltInCategory.OST_Ramps,
                    BuiltInCategory.OST_GenericModel,
                    BuiltInCategory.OST_MechanicalEquipment,
                    BuiltInCategory.OST_ElectricalEquipment,
                    BuiltInCategory.OST_PlumbingFixtures,
                    BuiltInCategory.OST_LightingFixtures,
                    BuiltInCategory.OST_DuctCurves,
                    BuiltInCategory.OST_PipeCurves,
                    BuiltInCategory.OST_CableTray,
                    BuiltInCategory.OST_Conduit
                };

                ElementMulticategoryFilter categoryFilter = new ElementMulticategoryFilter(categories);
                var filteredElements = collector.WherePasses(categoryFilter).ToElements();

                // Group elements by level
                var elementsByLevel = filteredElements
                    .GroupBy(e => doc.GetElement(e.LevelId)?.Name ?? "No Level")
                    .OrderBy(g => g.Key)
                    .ToDictionary(
                        g => g.Key,
                        g => g.GroupBy(e => e.Category?.Name ?? "Unknown")
                               .OrderBy(cg => cg.Key)
                               .ToDictionary(
                                   cg => cg.Key,
                                   cg => cg.Select(elem => CreateElementData(elem, doc)).ToList()
                               )
                    );

                // Get all levels
                var levels = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(l => l.Elevation)
                    .Select(l => new
                    {
                        Name = l.Name,
                        Id = l.Id.ToString(),
                        Elevation = l.Elevation
                    })
                    .ToList();

                // Get all views
                var views = new FilteredElementCollector(doc)
                    .OfClass(typeof(Autodesk.Revit.DB.View))
                    .Cast<Autodesk.Revit.DB.View>()
                    .Where(v => !v.IsTemplate)
                    .Select(v => new
                    {
                        Name = v.Name,
                        Id = v.Id.ToString(),
                        ViewType = v.ViewType.ToString(),
                        Scale = v.Scale
                    })
                    .ToList();

                // Build comprehensive project structure
                var projectData = new
                {
                    ProjectInfo = new
                    {
                        Title = doc.Title,
                        PathName = doc.PathName,
                        IsWorkshared = doc.IsWorkshared,
                        IsFamilyDocument = doc.IsFamilyDocument,
                        ExportDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                        RevitVersion = uiApp.Application.VersionNumber,
                        RevitBuild = uiApp.Application.VersionBuild
                    },
                    Statistics = new
                    {
                        TotalElements = filteredElements.Count,
                        ElementsByCategory = filteredElements
                            .GroupBy(e => e.Category?.Name ?? "Unknown")
                            .OrderBy(g => g.Key)
                            .ToDictionary(g => g.Key, g => g.Count()),
                        TotalLevels = levels.Count,
                        TotalViews = views.Count
                    },
                    Levels = levels,
                    Views = views,
                    Elements = elementsByLevel
                };

                // Serialize to JSON with indentation for readability
                string json = JsonConvert.SerializeObject(projectData, Formatting.Indented, new JsonSerializerSettings
                {
                    NullValueHandling = NullValueHandling.Ignore,
                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore
                });

                // Write to file
                File.WriteAllText(filePath, json);

                // Show success message
                Autodesk.Revit.UI.TaskDialog.Show(
                    "Export Complete",
                    $"Project exported successfully!\n\n" +
                    $"File: {filePath}\n" +
                    $"Elements exported: {filteredElements.Count}\n" +
                    $"File size: {new FileInfo(filePath).Length / 1024:N0} KB"
                );

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                Autodesk.Revit.UI.TaskDialog.Show("Export Error", $"An error occurred:\n\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        private object CreateElementData(Element element, Document doc)
        {
            return new
            {
                Id = element.Id.ToString(),
                UniqueId = element.UniqueId,
                Name = element.Name,
                Category = element.Category?.Name ?? "Unknown",
                Level = doc.GetElement(element.LevelId)?.Name ?? "N/A",
                TypeName = doc.GetElement(element.GetTypeId())?.Name ?? "N/A",
                Location = GetLocationInfo(element),
                Geometry = GetGeometryInfo(element),
                Properties = GetElementProperties(element, doc)
            };
        }

        private object GetLocationInfo(Element element)
        {
            if (element.Location == null)
                return null;

            try
            {
                if (element.Location is LocationPoint locationPoint)
                {
                    var point = locationPoint.Point;
                    double? rotation = null;
                    
                    try
                    {
                        // Not all LocationPoint elements support rotation
                        rotation = locationPoint.Rotation * (180.0 / Math.PI);
                    }
                    catch { }
                    
                    var result = new Dictionary<string, object>
                    {
                        { "Type", "Point" },
                        { "Point", new { X = point.X, Y = point.Y, Z = point.Z } }
                    };
                    
                    if (rotation.HasValue)
                        result["Rotation"] = rotation.Value;
                        
                    return result;
                }
                else if (element.Location is LocationCurve locationCurve)
                {
                    var curve = locationCurve.Curve;
                    var startPoint = curve.GetEndPoint(0);
                    var endPoint = curve.GetEndPoint(1);

                    return new
                    {
                        Type = "Curve",
                        StartPoint = new { X = startPoint.X, Y = startPoint.Y, Z = startPoint.Z },
                        EndPoint = new { X = endPoint.X, Y = endPoint.Y, Z = endPoint.Z },
                        Length = curve.Length,
                        CurveType = curve.GetType().Name
                    };
                }
            }
            catch { }

            return null;
        }

        private object GetGeometryInfo(Element element)
        {
            try
            {
                var bbox = element.get_BoundingBox(null);
                if (bbox != null)
                {
                    return new
                    {
                        BoundingBox = new
                        {
                            Min = new { X = bbox.Min.X, Y = bbox.Min.Y, Z = bbox.Min.Z },
                            Max = new { X = bbox.Max.X, Y = bbox.Max.Y, Z = bbox.Max.Z }
                        },
                        Width = bbox.Max.X - bbox.Min.X,
                        Depth = bbox.Max.Y - bbox.Min.Y,
                        Height = bbox.Max.Z - bbox.Min.Z
                    };
                }
            }
            catch { }

            return null;
        }

        private Dictionary<string, string> GetElementProperties(Element element, Document doc)
        {
            var properties = new Dictionary<string, string>();

            var paramNames = new[]
            {
                "Comments", "Mark", "Family", "Type",
                "Height", "Width", "Length", "Thickness",
                "Area", "Volume",
                "Head Height", "Sill Height",
                "Top Offset", "Base Offset", "Top Constraint", "Base Constraint",
                "Unconnected Height",
                "Fire Rating", "Function",
                "Structural Material", "Structural Usage"
            };

            foreach (var paramName in paramNames)
            {
                Parameter param = element.LookupParameter(paramName);
                if (param != null && param.HasValue)
                {
                    try
                    {
                        if (param.StorageType == StorageType.String)
                        {
                            var val = param.AsString();
                            if (!string.IsNullOrEmpty(val))
                                properties[paramName] = val;
                        }
                        else if (param.StorageType == StorageType.Double)
                        {
                            var val = param.AsDouble();
                            if (paramName.Contains("Area"))
                                properties[paramName] = $"{val:F2} sq ft";
                            else if (paramName.Contains("Volume"))
                                properties[paramName] = $"{val:F2} cu ft";
                            else
                                properties[paramName] = $"{val:F2} ft";
                        }
                        else if (param.StorageType == StorageType.Integer)
                            properties[paramName] = param.AsInteger().ToString();
                        else if (param.StorageType == StorageType.ElementId)
                        {
                            var elemId = param.AsElementId();
                            var linkedElem = doc.GetElement(elemId);
                            if (linkedElem != null)
                                properties[paramName] = linkedElem.Name;
                        }
                    }
                    catch { }
                }
            }

            // Add built-in parameters
            try
            {
                var areaParam = element.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                if (areaParam != null && areaParam.HasValue)
                    properties["ComputedArea"] = $"{areaParam.AsDouble():F2} sq ft";

                var volumeParam = element.get_Parameter(BuiltInParameter.HOST_VOLUME_COMPUTED);
                if (volumeParam != null && volumeParam.HasValue)
                    properties["ComputedVolume"] = $"{volumeParam.AsDouble():F2} cu ft";
            }
            catch { }

            return properties;
        }
    }
}

