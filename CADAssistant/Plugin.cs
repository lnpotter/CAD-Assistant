using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using System;
using System.Globalization;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

[assembly: CommandClass(typeof(CADAssistant.Plugin))]

namespace CADAssistant
{
    public class Plugin : IExtensionApplication
    {
        // Perplexity pplx-api settings
        private static readonly string ApiUrl = "https://api.perplexity.ai/chat/completions";
        private static readonly string Model = "sonar";

        public void Initialize()
        {
            var ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage(
                "\nCADAssistant loaded. Commands: AI_CHAT, AI_DRAW, AI_AUDIT_LAYERS\n"
            );
        }

        public void Terminate()
        {
        }

        // -------- Chat only (no drawing) --------
        [CommandMethod("AI_CHAT")]
        public void AiChat()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var ed = doc.Editor;

            var options = new PromptStringOptions("\nMessage for assistant: ")
            {
                AllowSpaces = true
            };

            var userInput = ed.GetString(options);
            if (userInput.Status != PromptStatus.OK)
                return;

            string context = GetDrawingContext(doc.Database);
            string answer = CallLlmText(
                $"Drawing context:\n{context}\n\nUser: {userInput.StringResult}"
            );

            ed.WriteMessage($"\n\n[CADAssistant]\n{answer}\n");
        }

        // -------- AI drawing + editing command --------
        [CommandMethod("AI_DRAW")]
        public void AiDraw()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            var options = new PromptStringOptions(
                "\nDescribe what you want to draw or edit (in English): "
            )
            {
                AllowSpaces = true
            };

            var userInput = ed.GetString(options);
            if (userInput.Status != PromptStatus.OK)
                return;

            string context = GetDrawingContext(db);

            string jsonPlan = CallLlmJsonPlan(
                $"Drawing context:\n{context}\n\nUser request: {userInput.StringResult}"
            );

            if (string.IsNullOrWhiteSpace(jsonPlan))
            {
                ed.WriteMessage("\n[CADAssistant] Empty plan from AI.\n");
                return;
            }

            ed.WriteMessage($"\n[CADAssistant] Plan JSON (raw):\n{jsonPlan}\n");

            jsonPlan = StripMarkdownCodeFence(jsonPlan);

            try
            {
                using var docJson = JsonDocument.Parse(jsonPlan);
                var root = docJson.RootElement;

                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var bt = (BlockTable)tr.GetObject(
                        db.BlockTableId,
                        OpenMode.ForRead
                    );
                    var btr = (BlockTableRecord)tr.GetObject(
                        bt[BlockTableRecord.ModelSpace],
                        OpenMode.ForWrite
                    );

                    if (root.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var action in root.EnumerateArray())
                        {
                            ExecuteAction(action, btr, tr);
                        }
                    }
                    else if (root.ValueKind == JsonValueKind.Object)
                    {
                        ExecuteAction(root, btr, tr);
                    }
                    else
                    {
                        ed.WriteMessage(
                            "\n[CADAssistant] Invalid JSON plan: root is not object or array.\n"
                        );
                    }

                    tr.Commit();
                }

                ed.WriteMessage("\n[CADAssistant] Plan executed.\n");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n[CADAssistant] Error executing plan: {ex.Message}\n");
            }
        }

        // -------- Simple audit example --------
        [CommandMethod("AI_AUDIT_LAYERS")]
        public void AiAuditLayers()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            var db = doc.Database;
            var ed = doc.Editor;

            string[] required = { "WIRE", "TITLE" };

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var lt = (LayerTable)tr.GetObject(
                    db.LayerTableId,
                    OpenMode.ForRead
                );

                foreach (var name in required)
                {
                    if (!lt.Has(name))
                    {
                        ed.WriteMessage($"\nMissing layer: {name}");
                        continue;
                    }

                    var ltr = (LayerTableRecord)tr.GetObject(
                        lt[name],
                        OpenMode.ForRead
                    );

                    if (ltr.IsOff)
                        ed.WriteMessage(
                            $"\nLayer {name} is OFF (recommended: turn it ON)."
                        );

                    if (ltr.IsFrozen)
                        ed.WriteMessage(
                            $"\nLayer {name} is FROZEN (recommended: thaw it)."
                        );

                    if (!ltr.IsPlottable)
                        ed.WriteMessage(
                            $"\nLayer {name} is not plottable (IsPlottable = false)."
                        );
                }

                tr.Commit();
            }

            ed.WriteMessage("\n\nLayer audit finished.\n");
        }

        // -------- Config helpers (API key) --------

        private static string GetApiKey()
        {
            try
            {
                string dllPath = System.Reflection.Assembly
                    .GetExecutingAssembly()
                    .Location;
                string baseDir = Path.GetDirectoryName(dllPath) ?? "";
                string configPath = Path.Combine(baseDir, "appsettings.local.json");

                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath, Encoding.UTF8);
                    using var doc = JsonDocument.Parse(json);
                    var root = doc.RootElement;

                    if (root.TryGetProperty("Perplexity", out var perplexity) &&
                        perplexity.ValueKind == JsonValueKind.Object &&
                        perplexity.TryGetProperty("ApiKey", out var keyProp) &&
                        keyProp.ValueKind == JsonValueKind.String)
                    {
                        string key = keyProp.GetString() ?? "";
                        if (!string.IsNullOrWhiteSpace(key))
                            return key;
                    }
                }
            }
            catch
            {
                // Ignore and fall back to environment variable
            }

            string? envKey = Environment.GetEnvironmentVariable("PPLX_API_KEY");
            if (!string.IsNullOrWhiteSpace(envKey))
                return envKey;

            throw new InvalidOperationException(
                "Perplexity API key not configured. Create appsettings.local.json with Perplexity.ApiKey next to CADAssistant.dll or set the PPLX_API_KEY environment variable."
            );
        }

        // -------- Helpers --------

        private static string GetDrawingContext(Database db)
        {
            return
                $"Insunits: {db.Insunits}\n"
                + $"Ltscale: {db.Ltscale}\n"
                + $"Cecolor (index): {db.Cecolor?.ColorIndex}\n";
        }

        private static string CallLlmText(string userText)
        {
            var responseText = CallLlmCommon(userText, false);
            return responseText ?? "(no content)";
        }

        private static string CallLlmJsonPlan(string userText)
        {
            var responseText = CallLlmCommon(userText, true);
            return responseText ?? "";
        }

        private static string CallLlmCommon(string userText, bool jsonPlan)
        {
            string apiKey;
            try
            {
                apiKey = GetApiKey();
            }
            catch (System.Exception ex)
            {
                return "[API KEY ERROR] " + ex.GetType().FullName + ": " + ex.Message;
            }

            using var http = new HttpClient();
            http.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", apiKey);
            http.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json")
            );

            string systemPrompt;

            if (jsonPlan)
            {
                systemPrompt =
                    "You are a CAD drawing planner for AutoCAD.\n" +
                    "Your job is to translate the user request into a JSON plan " +
                    "with drawing and editing actions, not instructions.\n\n" +
                    "Rules:\n" +
                    "- Always respond with pure JSON, no explanations, no Markdown.\n" +
                    "- Root must be an object or an array of objects.\n" +
                    "- Each object must have an 'action' field and parameters.\n" +
                    "- Supported actions:\n" +
                    "  1) 'line': {\"action\":\"line\",\"from\":[x1,y1],\"to\":[x2,y2],\"layer\":\"0\"}\n" +
                    "  2) 'polyline': {\"action\":\"polyline\",\"points\":[[x1,y1],[x2,y2],...],\"layer\":\"0\",\"closed\":false}\n" +
                    "  3) 'rectangle': {\"action\":\"rectangle\",\"base\":[x,y],\"width\":w,\"height\":h,\"rotation\":deg,\"layer\":\"0\"}\n" +
                    "  4) 'circle': {\"action\":\"circle\",\"center\":[x,y],\"radius\":r,\"layer\":\"0\"}\n" +
                    "  5) 'polygon': {\"action\":\"polygon\",\"center\":[x,y],\"radius\":r,\"sides\":n,\"rotation\":deg,\"layer\":\"0\"}\n" +
                    "  6) 'move': {\"action\":\"move\",\"filter\":{\"layer\":\"WIRE\"},\"offset\":[dx,dy]}\n" +
                    "  7) 'rotate': {\"action\":\"rotate\",\"filter\":{\"layer\":\"WIRE\"},\"base\":[x,y],\"angle\":deg}\n" +
                    "  8) 'scale': {\"action\":\"scale\",\"filter\":{\"layer\":\"WIRE\"},\"base\":[x,y],\"factor\":s}\n" +
                    "  9) 'erase': {\"action\":\"erase\",\"filter\":{\"layer\":\"WIRE\"}}\n" +
                    " 10) 'insert_block': {\"action\":\"insert_block\",\"name\":\"MOTOR\",\"position\":[x,y],\"scale\":1.0,\"rotation\":deg,\"layer\":\"0\",\"attributes\":{\"TAG\":\"M1\",\"DESC\":\"Motor\"}}\n" +
                    " 11) 'change_layer': {\"action\":\"change_layer\",\"filter\":{\"window\":[[x1,y1],[x2,y2]]},\"targetLayer\":\"CENTER\"}\n" +
                    " 12) 'change_properties': {\"action\":\"change_properties\",\"filter\":{\"layer\":\"0\"},\"colorIndex\":2,\"linetype\":\"CENTER\",\"linetypeScale\":0.5}\n" +
                    "- Filters can also specify a selection window or crossing window, for example:\n" +
                    "  {\"filter\":{\"window\":[[x1,y1],[x2,y2]]}} or {\"filter\":{\"crossing\":[[x1,y1],[x2,y2]]}}.\n" +
                    "- Coordinates are in drawing units.\n" +
                    "- Do not add comments or text outside JSON.";
            }
            else
            {
                systemPrompt =
                    "You are an assistant for AutoCAD Electrical. " +
                    "Answer briefly, in English, with clear step-by-step instructions " +
                    "for drawing, modifying, and auditing drawings.";
            }

            var payload = new
            {
                model = Model,
                messages = new object[]
                {
                    new { role = "system", content = systemPrompt },
                    new { role = "user", content = userText }
                },
                temperature = 0.1
            };

            var json = JsonSerializer.Serialize(payload);
            var response = http
                .PostAsync(ApiUrl, new StringContent(json, Encoding.UTF8, "application/json"))
                .GetAwaiter()
                .GetResult();

            var responseText = response.Content
                .ReadAsStringAsync()
                .GetAwaiter()
                .GetResult();

            if (!response.IsSuccessStatusCode)
                return $"HTTP error {response.StatusCode}: {responseText}";

            using var doc = JsonDocument.Parse(responseText);
            var content = doc.RootElement
                .GetProperty("choices")[0]
                .GetProperty("message")
                .GetProperty("content")
                .GetString();

            return content;
        }

        private static string StripMarkdownCodeFence(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return text;

            string trimmed = text.Trim();
            string fence = @"```";

            if (trimmed.StartsWith(fence + "json", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith(fence, StringComparison.OrdinalIgnoreCase))
            {
                int firstNewLine = trimmed.IndexOf('\n');
                if (firstNewLine >= 0)
                    trimmed = trimmed.Substring(firstNewLine + 1);
            }

            int lastFence = trimmed.LastIndexOf(fence, StringComparison.OrdinalIgnoreCase);
            if (lastFence >= 0)
            {
                trimmed = trimmed.Substring(0, lastFence);
            }

            return trimmed.Trim();
        }

        // -------- Execute actions from JSON --------

        private static void ExecuteAction(
            JsonElement actionElement,
            BlockTableRecord btr,
            Transaction tr
        )
        {
            if (!actionElement.TryGetProperty("action", out var actionProp))
                return;

            string action = actionProp.GetString() ?? "";
            action = action.ToLowerInvariant();

            string layerName = "0";
            if (actionElement.TryGetProperty("layer", out var layerProp) &&
                layerProp.ValueKind == JsonValueKind.String)
            {
                layerName = layerProp.GetString() ?? "0";
            }

            switch (action)
            {
                case "line":
                    CreateLine(actionElement, btr, tr, layerName);
                    break;

                case "polyline":
                    CreatePolyline(actionElement, btr, tr, layerName);
                    break;

                case "rectangle":
                    CreateRectangle(actionElement, btr, tr, layerName);
                    break;

                case "circle":
                    CreateCircle(actionElement, btr, tr, layerName);
                    break;

                case "polygon":
                    CreatePolygon(actionElement, btr, tr, layerName);
                    break;

                case "move":
                    EditMove(actionElement, btr, tr);
                    break;

                case "rotate":
                    EditRotate(actionElement, btr, tr);
                    break;

                case "scale":
                    EditScale(actionElement, btr, tr);
                    break;

                case "erase":
                    EditErase(actionElement, btr, tr);
                    break;

                case "insert_block":
                    CreateBlockInsert(actionElement, btr, tr, layerName);
                    break;

                case "change_layer":
                    EditChangeLayer(actionElement, btr, tr);
                    break;

                case "change_properties":
                    EditChangeProperties(actionElement, btr, tr);
                    break;

                default:
                    break;
            }
        }

        // -------- Drawing actions --------

        private static void CreateLine(
            JsonElement el,
            BlockTableRecord btr,
            Transaction tr,
            string layer
        )
        {
            var from = ReadPoint(el, "from", new Point2d(0, 0));
            var to = ReadPoint(el, "to", new Point2d(100, 0));

            var line = new Line(
                new Point3d(from.X, from.Y, 0),
                new Point3d(to.X, to.Y, 0)
            )
            {
                Layer = layer
            };

            btr.AppendEntity(line);
            tr.AddNewlyCreatedDBObject(line, true);
        }

        private static void CreatePolyline(
            JsonElement el,
            BlockTableRecord btr,
            Transaction tr,
            string layer
        )
        {
            var pline = new Polyline { Layer = layer };

            if (el.TryGetProperty("points", out var pts) &&
                pts.ValueKind == JsonValueKind.Array)
            {
                int idx = 0;
                foreach (var pt in pts.EnumerateArray())
                {
                    var p = ReadPointFromArray(pt, new Point2d(0, 0));
                    pline.AddVertexAt(idx++, p, 0, 0, 0);
                }
            }

            bool closed = el.TryGetProperty("closed", out var closedProp) &&
                          closedProp.ValueKind == JsonValueKind.True;

            pline.Closed = closed;

            btr.AppendEntity(pline);
            tr.AddNewlyCreatedDBObject(pline, true);
        }

        private static void CreateRectangle(
            JsonElement el,
            BlockTableRecord btr,
            Transaction tr,
            string layer
        )
        {
            var basePt = ReadPoint(el, "base", new Point2d(0, 0));
            double width = ReadDouble(el, "width", 100);
            double height = ReadDouble(el, "height", 100);
            double rotationDeg = ReadDouble(el, "rotation", 0);
            double rotRad = rotationDeg * Math.PI / 180.0;

            var p0 = new Point2d(basePt.X, basePt.Y);
            var p1 = new Point2d(basePt.X + width, basePt.Y);
            var p2 = new Point2d(basePt.X + width, basePt.Y + height);
            var p3 = new Point2d(basePt.X, basePt.Y + height);

            var pts = new[] { p0, p1, p2, p3 };

            var pline = new Polyline { Layer = layer };

            for (int i = 0; i < pts.Length; i++)
            {
                var p = RotateAround(pts[i], basePt, rotRad);
                pline.AddVertexAt(i, p, 0, 0, 0);
            }

            pline.Closed = true;

            btr.AppendEntity(pline);
            tr.AddNewlyCreatedDBObject(pline, true);
        }

        private static void CreateCircle(
            JsonElement el,
            BlockTableRecord btr,
            Transaction tr,
            string layer
        )
        {
            var center = ReadPoint(el, "center", new Point2d(0, 0));
            double radius = ReadDouble(el, "radius", 50);

            var circle = new Circle(
                new Point3d(center.X, center.Y, 0),
                Vector3d.ZAxis,
                radius
            )
            {
                Layer = layer
            };

            btr.AppendEntity(circle);
            tr.AddNewlyCreatedDBObject(circle, true);
        }

        private static void CreatePolygon(
            JsonElement el,
            BlockTableRecord btr,
            Transaction tr,
            string layer
        )
        {
            var center = ReadPoint(el, "center", new Point2d(0, 0));
            double radius = ReadDouble(el, "radius", 50);
            int sides = (int)ReadDouble(el, "sides", 6);
            double rotationDeg = ReadDouble(el, "rotation", 0);
            double rotRad = rotationDeg * Math.PI / 180.0;

            if (sides < 3)
                sides = 3;

            var pline = new Polyline { Layer = layer };

            for (int i = 0; i < sides; i++)
            {
                double angle = 2 * Math.PI * i / sides + rotRad;
                double x = center.X + radius * Math.Cos(angle);
                double y = center.Y + radius * Math.Sin(angle);
                pline.AddVertexAt(i, new Point2d(x, y), 0, 0, 0);
            }

            pline.Closed = true;

            btr.AppendEntity(pline);
            tr.AddNewlyCreatedDBObject(pline, true);
        }

        // -------- Block insertion --------

        private static void CreateBlockInsert(
            JsonElement el,
            BlockTableRecord btr,
            Transaction tr,
            string defaultLayer
        )
        {
            if (!el.TryGetProperty("name", out var nameProp) ||
                nameProp.ValueKind != JsonValueKind.String)
            {
                return;
            }

            string blockName = nameProp.GetString() ?? "";
            if (string.IsNullOrWhiteSpace(blockName))
                return;

            var pos2d = ReadPoint(el, "position", new Point2d(0, 0));
            var pos3d = new Point3d(pos2d.X, pos2d.Y, 0);

            double scale = ReadDouble(el, "scale", 1.0);
            double rotationDeg = ReadDouble(el, "rotation", 0.0);
            double rotationRad = rotationDeg * Math.PI / 180.0;

            string layer = defaultLayer;
            if (el.TryGetProperty("layer", out var layerProp) &&
                layerProp.ValueKind == JsonValueKind.String)
            {
                layer = layerProp.GetString() ?? defaultLayer;
            }

            Database db = btr.Database;
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

            if (!bt.Has(blockName))
            {
                return;
            }

            ObjectId bdefId = bt[blockName];

            var br = new BlockReference(pos3d, bdefId)
            {
                Layer = layer,
                ScaleFactors = new Scale3d(scale),
                Rotation = rotationRad
            };

            btr.AppendEntity(br);
            tr.AddNewlyCreatedDBObject(br, true);

            if (el.TryGetProperty("attributes", out var attrsEl) &&
                attrsEl.ValueKind == JsonValueKind.Object)
            {
                var btrDef = (BlockTableRecord)tr.GetObject(bdefId, OpenMode.ForRead);
                if (btrDef.HasAttributeDefinitions)
                {
                    foreach (ObjectId id in btrDef)
                    {
                        if (id.ObjectClass.DxfName != "ATTDEF")
                            continue;

                        var attDef = tr.GetObject(id, OpenMode.ForRead) as AttributeDefinition;
                        if (attDef == null || attDef.Constant)
                            continue;

                        string tag = attDef.Tag;
                        if (!attrsEl.TryGetProperty(tag, out var valProp))
                            continue;

                        if (valProp.ValueKind != JsonValueKind.String)
                            continue;

                        string value = valProp.GetString() ?? "";
                        var attRef = new AttributeReference();
                        attRef.SetAttributeFromBlock(attDef, br.BlockTransform);
                        attRef.TextString = value;

                        br.AttributeCollection.AppendAttribute(attRef);
                        tr.AddNewlyCreatedDBObject(attRef, true);
                    }
                }
            }
        }

        // -------- Editing actions --------

        private static void EditMove(JsonElement el, BlockTableRecord btr, Transaction tr)
        {
            var offset = ReadPoint(el, "offset", new Point2d(0, 0));
            var filter = el.TryGetProperty("filter", out var f) ? f : default;

            foreach (ObjectId id in btr)
            {
                var ent = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                if (ent == null) continue;
                if (!EntityMatchesFilter(ent, filter)) continue;

                ent.TransformBy(Matrix3d.Displacement(
                    new Vector3d(offset.X, offset.Y, 0)
                ));
            }
        }

        private static void EditRotate(JsonElement el, BlockTableRecord btr, Transaction tr)
        {
            var basePt = ReadPoint(el, "base", new Point2d(0, 0));
            double angleDeg = ReadDouble(el, "angle", 0);
            double angleRad = angleDeg * Math.PI / 180.0;
            var filter = el.TryGetProperty("filter", out var f) ? f : default;

            var center = new Point3d(basePt.X, basePt.Y, 0);

            foreach (ObjectId id in btr)
            {
                var ent = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                if (ent == null) continue;
                if (!EntityMatchesFilter(ent, filter)) continue;

                ent.TransformBy(Matrix3d.Rotation(
                    angleRad,
                    Vector3d.ZAxis,
                    center
                ));
            }
        }

        private static void EditScale(JsonElement el, BlockTableRecord btr, Transaction tr)
        {
            var basePt = ReadPoint(el, "base", new Point2d(0, 0));
            double factor = ReadDouble(el, "factor", 1.0);
            var filter = el.TryGetProperty("filter", out var f) ? f : default;

            var center = new Point3d(basePt.X, basePt.Y, 0);

            foreach (ObjectId id in btr)
            {
                var ent = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                if (ent == null) continue;
                if (!EntityMatchesFilter(ent, filter)) continue;

                ent.TransformBy(Matrix3d.Scaling(
                    factor,
                    center
                ));
            }
        }

        private static void EditErase(JsonElement el, BlockTableRecord btr, Transaction tr)
        {
            var filter = el.TryGetProperty("filter", out var f) ? f : default;

            foreach (ObjectId id in btr)
            {
                var ent = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                if (ent == null) continue;
                if (!EntityMatchesFilter(ent, filter)) continue;

                ent.Erase();
            }
        }

        private static void EditChangeLayer(JsonElement el, BlockTableRecord btr, Transaction tr)
        {
            if (!el.TryGetProperty("targetLayer", out var targetProp) ||
                targetProp.ValueKind != JsonValueKind.String)
            {
                return;
            }

            string targetLayer = targetProp.GetString() ?? "";
            if (string.IsNullOrWhiteSpace(targetLayer))
                return;

            var filter = el.TryGetProperty("filter", out var f) ? f : default;

            foreach (ObjectId id in btr)
            {
                var ent = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                if (ent == null) continue;
                if (!EntityMatchesFilter(ent, filter)) continue;

                ent.Layer = targetLayer;
            }
        }

        private static void EditChangeProperties(JsonElement el, BlockTableRecord btr, Transaction tr)
        {
            var filter = el.TryGetProperty("filter", out var f) ? f : default;

            bool hasColor = el.TryGetProperty("colorIndex", out var colorProp);
            short colorIndex = 256;
            if (hasColor)
            {
                try { colorIndex = (short)colorProp.GetInt32(); }
                catch
                {
                    if (colorProp.ValueKind == JsonValueKind.String &&
                        short.TryParse(colorProp.GetString(), out short c))
                        colorIndex = c;
                    else
                        hasColor = false;
                }
            }

            bool hasLinetype = el.TryGetProperty("linetype", out var ltProp) &&
                               ltProp.ValueKind == JsonValueKind.String;
            string linetype = hasLinetype ? (ltProp.GetString() ?? "") : "";

            bool hasLtScale = el.TryGetProperty("linetypeScale", out var ltsProp);
            double ltScale = 1.0;
            if (hasLtScale)
            {
                try { ltScale = ltsProp.GetDouble(); }
                catch
                {
                    if (ltsProp.ValueKind == JsonValueKind.String &&
                        double.TryParse(ltsProp.GetString(), NumberStyles.Float,
                            CultureInfo.InvariantCulture, out double v))
                        ltScale = v;
                    else
                        hasLtScale = false;
                }
            }

            foreach (ObjectId id in btr)
            {
                var ent = tr.GetObject(id, OpenMode.ForWrite, false) as Entity;
                if (ent == null) continue;
                if (!EntityMatchesFilter(ent, filter)) continue;

                if (hasColor)
                {
                    ent.ColorIndex = colorIndex;
                }

                if (hasLinetype && !string.IsNullOrWhiteSpace(linetype))
                {
                    ent.Linetype = linetype;
                }

                if (hasLtScale)
                {
                    ent.LinetypeScale = ltScale;
                }
            }
        }

        // -------- Filters (layer/type + window/crossing) --------

        private static bool EntityMatchesFilter(Entity ent, JsonElement filter)
        {
            if (filter.ValueKind == JsonValueKind.Undefined ||
                filter.ValueKind == JsonValueKind.Null)
                return true;

            // Filter by layer
            if (filter.TryGetProperty("layer", out var layerProp) &&
                layerProp.ValueKind == JsonValueKind.String)
            {
                string layer = layerProp.GetString() ?? "";
                if (!string.Equals(ent.Layer, layer, StringComparison.OrdinalIgnoreCase))
                    return false;
            }

            // Filter by type (line, circle, polyline, etc.)
            if (filter.TryGetProperty("type", out var typeProp) &&
                typeProp.ValueKind == JsonValueKind.String)
            {
                string type = typeProp.GetString() ?? "";
                string entType = ent.GetType().Name;
                if (!entType.Equals(type, StringComparison.OrdinalIgnoreCase))
                    return false;
            }

            // Window (inside)
            if (filter.TryGetProperty("window", out var winProp) &&
                winProp.ValueKind == JsonValueKind.Array)
            {
                if (!TryGetExtents(ent, out var ext))
                    return false;

                if (!WindowContains(winProp, ext))
                    return false;
            }

            // Crossing (intersects or inside)
            if (filter.TryGetProperty("crossing", out var crossProp) &&
                crossProp.ValueKind == JsonValueKind.Array)
            {
                if (!TryGetExtents(ent, out var ext))
                    return false;

                if (!WindowCrosses(crossProp, ext))
                    return false;
            }

            return true;
        }

        private static bool TryGetExtents(Entity ent, out Extents3d ext)
        {
            ext = new Extents3d();
            try
            {
                ext = ent.GeometricExtents;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool WindowContains(JsonElement windowEl, Extents3d ext)
        {
            var w = ReadWindow(windowEl, new Point2d(0, 0), new Point2d(0, 0));
            var minX = Math.Min(w.p1.X, w.p2.X);
            var maxX = Math.Max(w.p1.X, w.p2.X);
            var minY = Math.Min(w.p1.Y, w.p2.Y);
            var maxY = Math.Max(w.p1.Y, w.p2.Y);

            var emin = ext.MinPoint;
            var emax = ext.MaxPoint;

            return emin.X >= minX && emax.X <= maxX &&
                   emin.Y >= minY && emax.Y <= maxY;
        }

        private static bool WindowCrosses(JsonElement windowEl, Extents3d ext)
        {
            var w = ReadWindow(windowEl, new Point2d(0, 0), new Point2d(0, 0));
            var minX = Math.Min(w.p1.X, w.p2.X);
            var maxX = Math.Max(w.p1.X, w.p2.X);
            var minY = Math.Min(w.p1.Y, w.p2.Y);
            var maxY = Math.Max(w.p1.Y, w.p2.Y);

            var emin = ext.MinPoint;
            var emax = ext.MaxPoint;

            bool separated =
                emax.X < minX || emin.X > maxX ||
                emax.Y < minY || emin.Y > maxY;

            return !separated;
        }

        private static (Point2d p1, Point2d p2) ReadWindow(
            JsonElement arr,
            Point2d fallback1,
            Point2d fallback2
        )
        {
            Point2d p1 = fallback1;
            Point2d p2 = fallback2;

            try
            {
                int i = 0;
                foreach (var ptEl in arr.EnumerateArray())
                {
                    if (ptEl.ValueKind != JsonValueKind.Array)
                        continue;

                    var p = ReadPointFromArray(ptEl, new Point2d(0, 0));
                    if (i == 0) p1 = p;
                    else if (i == 1) { p2 = p; break; }
                    i++;
                }
            }
            catch
            {
            }

            return (p1, p2);
        }

        // -------- JSON reading helpers --------

        private static Point2d ReadPoint(
            JsonElement el,
            string name,
            Point2d fallback
        )
        {
            if (el.TryGetProperty(name, out var prop) &&
                prop.ValueKind == JsonValueKind.Array)
            {
                return ReadPointFromArray(prop, fallback);
            }

            return fallback;
        }

        private static Point2d ReadPointFromArray(
            JsonElement arr,
            Point2d fallback
        )
        {
            try
            {
                double x = fallback.X;
                double y = fallback.Y;

                int i = 0;
                foreach (var v in arr.EnumerateArray())
                {
                    if (i == 0)
                        x = v.GetDouble();
                    else if (i == 1)
                        y = v.GetDouble();
                    i++;
                }

                return new Point2d(x, y);
            }
            catch
            {
                return fallback;
            }
        }

        private static double ReadDouble(
            JsonElement el,
            string name,
            double fallback
        )
        {
            if (!el.TryGetProperty(name, out var prop))
                return fallback;

            try
            {
                return prop.GetDouble();
            }
            catch
            {
                if (prop.ValueKind == JsonValueKind.String &&
                    double.TryParse(
                        prop.GetString(),
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        out double val
                    ))
                {
                    return val;
                }

                return fallback;
            }
        }

        private static Point2d RotateAround(
            Point2d p,
            Point2d center,
            double angleRad
        )
        {
            double cos = Math.Cos(angleRad);
            double sin = Math.Sin(angleRad);

            double dx = p.X - center.X;
            double dy = p.Y - center.Y;

            double x = center.X + dx * cos - dy * sin;
            double y = center.Y + dx * sin + dy * cos;

            return new Point2d(x, y);
        }
    }
}