# Claude Code Instructions for Ollama Excel UDF

## Project Overview

This is a VBA Excel add-in that provides AI functions for Excel spreadsheets. It integrates with Ollama (local), OpenAI, Perplexity, Gemini, and other OpenAI-compatible APIs.

## Critical: VBA File Requirements

### Line Endings

**All VBA files MUST use Windows-style CRLF line endings.**

| File Type | Extension | Line Ending | Notes |
|-----------|-----------|-------------|-------|
| Module | `.bas` | CRLF | Standard VBA module |
| Form | `.frm` | CRLF | **Critical** - mixed endings will prevent import |
| Class | `.cls` | CRLF | Class module |

After editing any VBA file, verify line endings:
```bash
file src/filename.bas
# Should show: "ASCII text, with CRLF line terminators"
```

If mixed or LF-only, convert:
```bash
unix2dos src/filename.bas
```

### Form Files (.frm)

Form files have two parts:
1. `.frm` - Text file with form header + VBA code
2. `.frx` - Binary file with visual control definitions

**You cannot edit `.frx` files programmatically.** Visual controls (buttons, checkboxes, textboxes) must be added manually in the VBA editor.

When adding new form controls:
1. Document the control properties in code comments
2. Add the control reference in `UserForm_Initialize`
3. Instruct user to manually add the control in VBA editor

Example:
```vba
' Requires control: chkPromptRow (CheckBox)
' - Name: chkPromptRow
' - Caption: "Row 1 contains column-specific prompts"
Private Sub UserForm_Initialize()
    chkPromptRow.Value = False  ' Will error if control not added
End Sub
```

## Project Structure

```
ollama-excel-udf/
├── add-in/
│   └── OllamaLLM.xlam          # Compiled add-in (binary, do not edit)
├── src/
│   ├── JsonConverter.bas       # Third-party JSON library (do not modify)
│   ├── modAI_Function.bas      # All AI UDF implementations
│   ├── modAI_Tooltips.bas      # Function registration and IntelliSense
│   ├── modAI_Bulk.bas          # Bulk fill macro logic
│   └── frmAIBulk.frm           # UserForm code (needs .frx for controls)
├── assets/                     # Demo GIFs and images
├── README.md                   # User documentation
├── AGENTS.md                   # Project context for AI agents
├── BUILD_INSTRUCTIONS.md       # How to build the add-in
└── CLAUDE.md                   # This file
```

## Key Files

### modAI_Function.bas

Contains all public AI functions:

| Function | Signature |
|----------|-----------|
| `AI` | `(prompt, [model], [temp], [max_tokens], [endpoint], [api_key])` |
| `AI_SEARCH` | Same as AI, uses `[search]` INI section |
| `AI_EXTRACT` | `(text, field, ...)` |
| `AI_CLASSIFY` | `(text, categories, ...)` - categories can be String or Range |
| `AI_TRANSLATE` | `(text, targetLang, ...)` |
| `AI_SUMMARIZE` | `(text, [maxWords], ...)` - default 50 words |
| `AI_SENTIMENT` | `(text, ...)` - returns Positive/Negative/Neutral |
| `AI_FIX` | `(text, [rules], ...)` |

**Internal functions:**
- `AI_Core` - Private function that handles HTTP requests, called by all AI_* functions
- `ParseCategories` - Converts Range or String to comma-separated category list
- `BuildChatPayload` - Creates OpenAI-compatible JSON payload
- `BuildGeminiPayload` - Creates Gemini API payload
- `BuildResponsesPayload` - Creates OpenAI Responses API payload

### modAI_Bulk.bas

Bulk fill logic for processing tables:

**Key functions:**
- `RunBulkFill(ui As frmAIBulk)` - Main entry point
- `BuildPrompt(columnPrompt, globalPrompt, headerText, contextText)` - Constructs final prompt
- `BuildRowContext(headerRow, dataRow, inputCols)` - Gathers input column data

**Modes:**
1. Standard: Row 1 = headers, Row 2+ = data
2. Prompt Row: Row 1 = per-column prompts, Row 2 = headers, Row 3+ = data

### modAI_Tooltips.bas

Registers all functions with Excel for IntelliSense:

- Each function needs a `RegisterXxx` private function
- Called from `Install_AI_Tooltips` on workbook open
- Uses `Application.MacroOptions` for registration

### frmAIBulk.frm

UserForm for bulk operations:

**Required controls (must exist in .frx):**
- `txtPrompt` - TextBox for global prompt
- `chkSearch` - CheckBox for search mode
- `chkPromptRow` - CheckBox for prompt row mode
- `btnRun` - CommandButton to start
- `btnClose` - CommandButton to close/stop
- `lblStatus` - Label for status messages

**Public methods exposed to modAI_Bulk:**
- `PromptText() As String`
- `IsSearchMode() As Boolean`
- `HasPromptRow() As Boolean`
- `UpdateStatus(message As String)`
- `Cancelled As Boolean` (public property)

## Configuration System

Settings stored in `%APPDATA%\OllamaLLM\config.ini`:

```ini
[ai]
model = qwen3:30b-a3b-instruct-2507-q8_0
endpoint = http://localhost:11434
api_key = 
temperature = 0.2
max_tokens = 512
system = You are a helpful assistant...

[search]
model = sonar-pro
endpoint = https://api.perplexity.ai
api_key = YOUR_KEY
...
```

**INI functions in modAI_Function.bas:**
- `GetIniPath()` - Returns config file path
- `ReadIniValue(section, key, default)` - Reads value
- `WriteIniDefault(section, key, value)` - Writes if not exists
- `EnsureIniDefaults()` - Creates default config on first run
- `ResolveIniString/Double/Long()` - Gets value with fallback

## Adding New Functions

### 1. Add the UDF to modAI_Function.bas

```vba
Public Function AI_NEWFUNCTION(text As String, _
                               Optional param As String = "", _
                               Optional model As String = "", _
                               Optional temperature As Variant, _
                               Optional max_tokens As Variant, _
                               Optional endpoint As String = "", _
                               Optional api_key As String = "") As String
    Dim systemPrompt As String
    Dim resolvedModel As String
    ' ... resolve parameters from INI ...
    
    systemPrompt = "Your instruction here. " & param
    
    AI_NEWFUNCTION = AI_Core(text, systemPrompt, resolvedModel, ...)
End Function
```

### 2. Add tooltip registration to modAI_Tooltips.bas

```vba
Private Function RegisterNewFunction(ByVal macroName As String) As Boolean
    On Error GoTo Fail
    Application.MacroOptions _
        Macro:=macroName, _
        Description:="Description for function wizard.", _
        Category:="AI Helpers", _
        ArgumentDescriptions:=Array( _
            "text (required): The input text.", _
            "param (optional): Additional parameter." _
        )
    RegisterNewFunction = True
    Exit Function
Fail:
    RegisterNewFunction = False
End Function
```

### 3. Register in Install_AI_Tooltips

```vba
If Not RegisterNewFunction(ThisWorkbook.Name & "!AI_NEWFUNCTION") Then errs = errs & ...
If Not RegisterNewFunction("AI_NEWFUNCTION") Then errs = errs & ...
```

## Testing Checklist

After making changes:

1. [ ] Verify CRLF line endings on all modified .bas/.frm files
2. [ ] Import files into Excel VBA editor
3. [ ] Enable Microsoft Scripting Runtime reference
4. [ ] Test each modified function in a cell
5. [ ] Test bulk fill in both modes (standard and prompt row)
6. [ ] Verify tooltips appear in function wizard

## Common Issues

### "Can't import form" or form imports as module
- **Cause**: Mixed or LF-only line endings
- **Fix**: Convert to CRLF with `unix2dos`

### "Variable not defined" for form controls
- **Cause**: Control not added to form designer
- **Fix**: User must manually add control in VBA editor

### "Compile error: User-defined type not defined"
- **Cause**: Missing Microsoft Scripting Runtime reference
- **Fix**: Tools → References → Enable "Microsoft Scripting Runtime"

### Function not appearing in Excel
- **Cause**: Tooltips not registered
- **Fix**: Run `Install_AI_Tooltips` or reload add-in

## API Endpoints

The add-in supports multiple API formats:

| Provider | Endpoint Format | Payload Builder |
|----------|-----------------|-----------------|
| Ollama/OpenAI | `/v1/chat/completions` | `BuildChatPayload` |
| Perplexity | `/chat/completions` | `BuildChatPayload` |
| Gemini | `/models/{model}:generateContent` | `BuildGeminiPayload` |
| OpenAI Responses | `/v1/responses` | `BuildResponsesPayload` |

Endpoint normalization happens in:
- `NormalizeEndpoint()` - Standard OpenAI format
- `NormalizeGeminiEndpoint()` - Gemini format with API key in URL
- `NormalizeResponsesEndpoint()` - OpenAI Responses API

## Dependencies

- **WinHTTP** - Built into Windows, used for HTTP requests
- **Microsoft Scripting Runtime** - For `Scripting.Dictionary` objects
- **VBA-JSON** - JsonConverter.bas for JSON parsing (do not modify)
