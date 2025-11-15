using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using revit_mcp_plugin.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace revit_mcp_plugin.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ShowCurrentViewElementsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;
                View activeView = doc.ActiveView;

                // Get elements in the current view
                var collector = new FilteredElementCollector(doc, activeView.Id)
                    .WhereElementIsNotElementType();

                // Filter to common categories
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
                    BuiltInCategory.OST_Rooms
                };

                ElementMulticategoryFilter categoryFilter = new ElementMulticategoryFilter(categories);
                var filteredElements = collector.WherePasses(categoryFilter).ToElements();

                // Filter out hidden elements
                var visibleElements = filteredElements.Where(e => !e.IsHidden(activeView)).ToList();

                // Build result object with comprehensive element data
                var result = new
                {
                    ViewName = activeView.Name,
                    ViewId = activeView.Id.ToString(),
                    TotalElementsInView = new FilteredElementCollector(doc, activeView.Id).WhereElementIsNotElementType().GetElementCount(),
                    FilteredElementCount = visibleElements.Count,
                    Elements = visibleElements.Select(e => new
                    {
                        Id = e.Id.ToString(),
                        UniqueId = e.UniqueId,
                        Name = e.Name,
                        Category = e.Category?.Name ?? "Unknown",
                        Level = doc.GetElement(e.LevelId)?.Name ?? "N/A",
                        TypeName = doc.GetElement(e.GetTypeId())?.Name ?? "N/A",
                        Location = GetLocationInfo(e),
                        Geometry = GetGeometryInfo(e),
                        Properties = GetElementProperties(e, doc)
                    }).ToList()
                };

                // Convert to JSON
                string jsonResult = JsonConvert.SerializeObject(result, Formatting.Indented);

                // Show results in a window
                var resultsWindow = new ResultsWindow
                {
                    Title = $"Current View Elements - {activeView.Name}",
                    ResultText = jsonResult
                };
                resultsWindow.ShowDialog();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("Error", $"An error occurred:\n\n{ex.Message}\n\nStack Trace:\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        private object GetLocationInfo(Element element)
        {
            if (element.Location == null)
                return null;

            if (element.Location is LocationPoint locationPoint)
            {
                var point = locationPoint.Point;
                return new
                {
                    Type = "Point",
                    Point = new { X = point.X, Y = point.Y, Z = point.Z },
                    Rotation = locationPoint.Rotation * (180.0 / Math.PI) // Convert to degrees
                };
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

            // Get ALL parameters from the element
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
                            // Check if it's an area or volume parameter
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

