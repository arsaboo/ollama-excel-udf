# Project Context for Ollama Excel UDF

## Project Overview
This is an Excel add-in (UDF - User Defined Function) that integrates with AI providers (Ollama, OpenAI, Perplexity, Gemini) to provide AI capabilities directly within Excel spreadsheets. The add-in provides 8 AI functions for common tasks like extraction, classification, translation, summarization, and sentiment analysis. It also includes a bulk fill UserForm for processing entire tables with AI.

### Key Technologies
- **Language**: VBA (Visual Basic for Applications)
- **Platform**: Excel for Windows
- **API Integration**: OpenAI-compatible `/v1/chat/completions` endpoint, Gemini API, OpenAI Responses API
- **Dependencies**: 
  - WinHTTP for HTTP requests
  - Microsoft Scripting Runtime for Dictionary objects
  - VBA-JSON library for JSON parsing

### Architecture
The project consists of:
1. `modAI_Function.bas` - Contains all AI UDF implementations (AI, AI_SEARCH, AI_EXTRACT, AI_CLASSIFY, AI_TRANSLATE, AI_SUMMARIZE, AI_SENTIMENT, AI_FIX)
2. `modAI_Tooltips.bas` - Handles function tooltips and IntelliSense registration
3. `modAI_Bulk.bas` - Runs bulk fill logic for header-aware tables with per-column prompt support
4. `frmAIBulk.frm` - UserForm for bulk fill operations
5. `JsonConverter.bas` - Third-party JSON parsing library
6. `OllamaLLM.xlam` - Compiled Excel add-in file

## Project Structure
```
ollama-excel-udf/
├── add-in/
│   └── OllamaLLM.xlam          # Compiled Excel add-in
├── dev-workbook/
│   └── README.md               # Instructions for development workbook
├── src/
│   ├── JsonConverter.bas       # JSON parsing library (VBA-JSON)
│   ├── modAI_Function.bas      # All AI function implementations
│   ├── modAI_Tooltips.bas      # Function tooltips and registration
│   ├── modAI_Bulk.bas          # Bulk fill macro logic
│   └── frmAIBulk.frm           # UserForm module
├── assets/                     # GIF demos and images
├── BUILD_INSTRUCTIONS.md       # Detailed build instructions
├── README.md                   # User-facing documentation
├── LICENSE                     # MIT License
├── .gitignore                  # Git ignore rules
└── AGENTS.md                   # This file
```

## Available Functions

### Core Functions
| Function | Purpose |
|----------|---------|
| `AI()` | General-purpose AI prompt |
| `AI_SEARCH()` | AI with web search (Perplexity/Gemini) |

### Specialized Functions
| Function | Purpose |
|----------|---------|
| `AI_EXTRACT(text, field)` | Extract specific data (email, phone, date, etc.) |
| `AI_CLASSIFY(text, categories)` | Classify into categories (string or range) |
| `AI_TRANSLATE(text, targetLang)` | Translate to any language |
| `AI_SUMMARIZE(text, [maxWords])` | Summarize to word count (default: 50) |
| `AI_SENTIMENT(text)` | Returns: Positive, Negative, or Neutral |
| `AI_FIX(text, [rules])` | Fix grammar, spelling, formatting |

## Bulk Fill Features
- **Standard Mode**: Row 1 = headers, Row 2+ = data
- **Per-Column Prompt Mode**: Row 1 = custom prompts per column, Row 2 = headers, Row 3+ = data
- **Hotkey**: `Ctrl+Shift+A` opens the form
- **Auto-detection**: Input columns (with data) vs output columns (empty with headers)

## Configuration
Settings stored in `%APPDATA%\OllamaLLM\config.ini`:
- `[ai]` section for standard AI functions
- `[search]` section for AI_SEARCH function
- Supports: model, endpoint, api_key, temperature, max_tokens, system prompt

## Supported Providers
- **Ollama** (local/self-hosted)
- **OpenAI** (GPT models)
- **Perplexity** (Sonar models with search)
- **Google Gemini** (with Google Search)
- **OpenRouter** (multi-provider gateway)
- Any OpenAI-compatible API

## Development Workflow
1. **Editing Source Code**: Modify `.bas`/`.frm` files in `src/`
2. **Building the Add-in**: 
   - Import files into Excel VBA editor
   - Enable Microsoft Scripting Runtime reference
   - Add `chkPromptRow` checkbox to `frmAIBulk` form
   - Save as `.xlam` in `add-in/`
3. **Testing**: Install add-in and test all 8 functions + bulk fill modes

## Key Implementation Details

### AI_Core Function
All AI functions call the private `AI_Core()` function which handles:
- HTTP request setup (WinHTTP)
- Payload building (JSON)
- Response parsing
- Error handling

### ParseCategories Helper
`AI_CLASSIFY` accepts both:
- Comma-separated string: `"Tech, Sports, Politics"`
- Cell range: `$A$1:$A$10`

The `ParseCategories()` function handles both input types.

### Bulk Fill Prompt Resolution
For each output cell:
1. Check if column has a prompt in Row 1 (when prompt row mode enabled)
2. If yes, use column-specific prompt
3. If no, fall back to global prompt from form
4. Append header name and row context

## Requirements
- Excel for Windows (uses WinHTTP)
- For local use: Ollama server with model pulled
- For cloud use: API key for respective provider

## Security Notes
- Local models (Ollama): Data never leaves machine
- Cloud providers: Data sent to provider API
- No telemetry in add-in
- Consider signing add-in to reduce macro warnings

## License
MIT License - see LICENSE file for details
