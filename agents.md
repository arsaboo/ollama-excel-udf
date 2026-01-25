# Project Context for Ollama Excel UDF

## Project Overview
This is an Excel add-in (UDF - User Defined Function) that integrates with Ollama AI servers to provide AI capabilities directly within Excel spreadsheets. The main function is `=AI()` which allows users to send prompts to an Ollama server and receive concise, cell-friendly answers. The add-in also includes a macro-driven bulk fill UserForm for header-aware table fills.

### Key Technologies
- **Language**: VBA (Visual Basic for Applications)
- **Platform**: Excel for Windows
- **API Integration**: Ollama server via OpenAI-compatible `/v1/chat/completions` endpoint
- **Dependencies**: 
  - WinHTTP for HTTP requests
  - Microsoft Scripting Runtime for Dictionary objects
  - VBA-JSON library for JSON parsing

### Architecture
The project consists of:
1. `modAI_Function.bas` - Contains the main AI UDF implementation
2. `modAI_Tooltips.bas` - Handles function tooltips and registration
3. `modAI_Bulk.bas` - Runs bulk fill logic for header-aware tables
4. `frmAIBulk.frm` - UserForm for bulk fill prompts
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
│   ├── JsonConverter.bas       # JSON parsing library
│   ├── modAI_Function.bas      # Main AI function implementation
│   ├── modAI_Tooltips.bas      # Function tooltips and registration
│   ├── modAI_Bulk.bas          # Bulk fill macro logic
│   └── frmAIBulk.frm           # UserForm module
├── BUILD_INSTRUCTIONS.md       # Detailed instructions for building the add-in
├── README.md                   # Project documentation
├── LICENSE                     # MIT License
├── .gitignore                  # Git ignore rules
└── QWEN.md                     # This file
```

## Key Features
1. **AI Function**: `=AI(prompt, [model], [temperature], [max_tokens], [system], [endpoint])`
2. **AI Search Function**: `=AI_SEARCH(prompt, [model], [temperature], [max_tokens], [system], [endpoint])`
3. **Bulk Fill UserForm**: `Show_AI_Form` macro fills tables based on column headers and row context
   - Hotkey: `Ctrl+Shift+A` opens the form after the add-in loads
4. **Configurable Parameters**:
   - Prompt (required)
   - Model selection (default: qwen3:30b-a3b-instruct-2507-q8_0)
   - Temperature control (default: 0.2)
   - Max tokens (default: 512)
   - System prompt (optional)
   - Custom endpoint (default: http://192.168.2.162:11434/v1/chat/completions)

## Development Workflow
1. **Editing Source Code**: Modify the .bas/.frm files in the `src/` directory
2. **Building the Add-in**: 
   - Import the .bas/.frm files into Excel VBA editor
   - Enable Microsoft Scripting Runtime reference
   - Save as .xlam file in the `add-in/` directory
3. **Testing**: Install the add-in in Excel and test both the AI function and bulk fill form

## Installation Process
1. Locate the add-in file: `/add-in/OllamaLLM.xlam`
2. In Excel: `File → Options → Add-ins → Manage: Excel Add-ins → Go… → Browse…`
3. Select the `OllamaLLM.xlam` file and ensure it's checked
4. The AI function will be available in Excel

## Requirements
- Excel for Windows (uses WinHTTP)
- Accessible Ollama server
- Required model pulled on the Ollama server

## Development Considerations
- All development happens in VBA within Excel
- The add-in uses WinHTTP for network requests
- JSON responses are parsed using the VBA-JSON library
- Function tooltips are registered automatically when the workbook opens
- Error handling is implemented for HTTP and VBA errors

## Security Notes
- Communicates directly with Ollama host (no API key)
- If exposing beyond LAN, protect the host appropriately
- Consider signing the add-in to reduce macro warnings

## License
MIT License - see LICENSE file for details
