# CAD-Assistant

CAD-Assistant is an experimental AutoCAD Electrical plugin that connects to the Perplexity AI API to provide an AI-powered assistant inside AutoCAD.

The plugin adds commands that let you chat with an AI model, automatically plan drawing actions in JSON, and execute those actions directly in the current drawing.

---

## Features

- **AI_CHAT**  
  Free-form chat with the AI assistant about drawing, editing, and auditing tasks in AutoCAD Electrical.

- **AI_DRAW**  
  - Accepts a natural language request (in English).  
  - Sends context + request to the Perplexity API.  
  - Receives a JSON "plan" describing drawing/editing actions.  
  - Executes the plan in Model Space (lines, polylines, rectangles, circles, polygons, block inserts, moves, scales, rotates, erases, layer/property edits, etc.).

- **AI_AUDIT_LAYERS**  
  Simple example of an automated audit that checks for required layers (e.g., `WIRE`, `TITLE`) and reports visibility/plot settings.

---

## JSON Action Model

The AI is instructed to return pure JSON with actions such as:

- `line`, `polyline`, `rectangle`, `circle`, `polygon`  
- `move`, `rotate`, `scale`, `erase`  
- `insert_block` (with attributes)  
- `change_layer` (mass layer reassignment by filter)  
- `change_properties` (color index, linetype, linetype scale by filter)

Each action may include a `filter` object to select entities by:

- `layer`  
- `type` (e.g. `Line`, `Polyline`, `Circle`)  
- `window`: `{"window":[[x1,y1],[x2,y2]]}` for entities fully inside a rectangular window  
- `crossing`: `{"crossing":[[x1,y1],[x2,y2]]}` for entities that touch or lie inside the window

The C# plugin parses this JSON and applies the requested transformations using AutoCAD .NET APIs.

---

## Requirements

- Windows with AutoCAD Electrical (tested with recent versions).  
- .NET and Visual Studio 2022 (or compatible) to build the project.  
- A Perplexity API key with access to the `sonar` model.

---

## Configuration

Inside `Plugin.cs` you will find:

private static readonly string ApiUrl = "https://api.perplexity.ai/chat/completions";
private static readonly string ApiKey = "YOUR_API_KEY_HERE";
private static readonly string Model = "sonar";


Before using the plugin:

1. Replace `YOUR_API_KEY_HERE` with your own Perplexity API key.  
2. Rebuild the project and reload the plugin in AutoCAD.

---

## Building

1. Open the solution in Visual Studio 2022.  
2. Make sure the target .NET Framework matches your AutoCAD version.  
3. Build the project (**Build → Rebuild Solution**).  
4. The resulting `.dll` will be created in `bin/Debug` or `bin/Release`.

---

## Installing in AutoCAD

1. Copy the compiled `.dll` to a folder of your choice (for example, `C:\CADAssistant`).  
2. In AutoCAD, run the `NETLOAD` command and select the plugin `.dll`.  
3. After loading, the command line should display the available commands:
   - `AI_CHAT`  
   - `AI_DRAW`  
   - `AI_AUDIT_LAYERS`

You can then start experimenting by describing operations in English, such as:

- `Move all entities inside window from (0,0) to (200,100) to layer CENTER.`  
- `Set all entities on layer 0 to color index 2, linetype CENTER, linetype scale 0.5.`  

---

## Roadmap

Planned improvements include:

- Mirror and array actions (e.g. `mirror`, `array_linear`).  
- Filters by geometric length/area (e.g. erase short segments, small hatches).  
- Discipline-specific presets (electrical, mechanical, architectural).  

Contributions, ideas, and issues are welcome via GitHub.