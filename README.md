# AI Functions for Excel

**Bring the power of AI directly into your spreadsheets.** Extract data, classify text, translate content, analyze sentiment, and more — all with simple Excel formulas.

Works with **Ollama** (local/self-hosted), **OpenAI**, **Perplexity**, **Gemini**, and any OpenAI-compatible API.

![Capitals Example](assets/Capitals.gif)

---

## Why Use This Add-in?

| Feature | Benefit |
|---------|---------|
| **8 AI Functions** | Purpose-built formulas for common tasks — no prompt engineering required |
| **Bulk Fill Tool** | Process hundreds of rows with a single click using `Ctrl+Shift+A` |
| **Per-Column Prompts** | Different AI instructions for each column in your table |
| **Works Offline** | Use local Ollama models — your data never leaves your machine |
| **Cloud Compatible** | Switch to OpenAI, Perplexity, or Gemini when you need more power |
| **Zero Cost** | Free and open source. No subscriptions, no per-token fees with local models |

---

## Available Functions

### Core Functions

| Function | What It Does | Example |
|----------|--------------|---------|
| `AI()` | General-purpose AI prompt | `=AI("What is the capital of France?")` |
| `AI_SEARCH()` | AI with web search (Perplexity/Gemini) | `=AI_SEARCH("Latest GDP of Japan")` |

### Specialized Functions

| Function | What It Does | Example |
|----------|--------------|---------|
| `AI_EXTRACT()` | Pull specific data from text | `=AI_EXTRACT(A1, "email")` |
| `AI_CLASSIFY()` | Categorize text | `=AI_CLASSIFY(A1, "Tech, Sports, Politics")` |
| `AI_TRANSLATE()` | Translate to any language | `=AI_TRANSLATE(A1, "Spanish")` |
| `AI_SUMMARIZE()` | Condense long text | `=AI_SUMMARIZE(A1, 50)` |
| `AI_SENTIMENT()` | Analyze sentiment | `=AI_SENTIMENT(A1)` → Positive/Negative/Neutral |
| `AI_FIX()` | Fix grammar and spelling | `=AI_FIX(A1, "formal tone")` |

---

## Quick Examples

### Extract Data from Text
```excel
=AI_EXTRACT(A1, "phone number")
=AI_EXTRACT(A1, "company name")
=AI_EXTRACT(A1, "date")
```

### Classify Customer Feedback
```excel
=AI_CLASSIFY(A1, "Bug Report, Feature Request, Question, Complaint")
=AI_CLASSIFY(A1, $D$1:$D$10)   ' Use a range of categories
```

### Translate Product Descriptions
```excel
=AI_TRANSLATE(A1, "German")
=AI_TRANSLATE(A1, "Japanese")
```

### Analyze Reviews
```excel
=AI_SENTIMENT(A1)   ' Returns: Positive, Negative, or Neutral
```

### Clean Up Text
```excel
=AI_FIX(A1)                        ' Fix grammar and spelling
=AI_FIX(A1, "British English")     ' Apply specific rules
=AI_FIX(A1, "formal business tone")
```

### Summarize Long Content
```excel
=AI_SUMMARIZE(A1)       ' Default: 50 words
=AI_SUMMARIZE(A1, 25)   ' Custom word limit
```

---

## Bulk Fill: Process Tables with AI

The **Bulk Fill** tool lets you fill entire tables without writing formulas in every cell.

**How it works:**
1. Select any cell in your table
2. Press `Ctrl+Shift+A` to open the AI Agent form
3. Enter a global prompt and click **Run**

The tool automatically:
- Uses your **column headers** as output requirements
- Uses **populated columns** as context for each row
- Fills all empty columns with AI-generated content

**Global Prompt Examples (pick one that matches your table):**

**Simple & Universal (works for any table):**
```text
Fill in this table using the column headers as guidance.
```

**Product Catalog:**
```text
Analyze the product information and fill in appropriate details.
```

**Data Enrichment:**
```text
Extract and structure the missing information from the provided data.
```

**Content Generation:**
```text
Generate appropriate content for each column based on the row data.
```

**Categorization:**
```text
Review the data and categorize each row into the appropriate columns.
```

### Per-Column Prompts Mode (when you need different AI tasks per column)
| Product Title      | Description        | Category    | Meta Title  |
|--------------------|--------------------| (empty)     | (empty)     |
| BIANCHI BIKE 16"   | Introducing the new... | ← AI fills  | ← AI fills  |
| GIANT BIKE 700C    | The all-new model... | ← AI fills  | ← AI fills  |
```

**Global prompt:** `"Fill in this table using the column headers as guidance."`
> ✅ *Use this when your headers are self-explanatory*

### Per-Column Prompts Mode (Prompts in Row 1)

For advanced control where each column needs specific instructions.

```
| (empty)            | (empty)            | "Categorize as Bike/Accessory/Clothing" | "Write SEO title under 60 chars" |
| Product Title      | Description        | Category                                 | Meta Title                        |
| BIANCHI BIKE 16"   | Introducing the new... | ← Uses column prompt                     | ← Uses column prompt              |
| GIANT BIKE 700C    | The all-new model... | ← Uses column prompt                     | ← Uses column prompt              |
```

**Simple Global Prompt (works for any table):**
```
Fill in this table using the column headers as guidance.
```

Each column can have its own custom instruction. Empty prompt cells fall back to the global prompt.

> ✅ *Perfect when you need different AI tasks for each column*

---

## Quick Start (3 Minutes to AI in Excel)

### 1. Install & Configure
1. **Download** `OllamaLLM.xlam` from [Releases](../../releases)
2. **Install** in Excel: `File → Options → Add-ins → Manage: Excel Add-ins → Go → Browse`
3. **Configure** AI provider (if not using local Ollama):
   - Press `Alt+F8` → Run `Open_AI_Config`
   - Edit `%APPDATA%\OllamaLLM\config.ini`

### 2. Start Using AI Functions

**Try these in any Excel cell:**
```excel
=AI("What is 15% of 847?")              ' → 127.05
=AI_EXTRACT("Call me at 555-1234", "phone")  ' → 555-1234
=AI_TRANSLATE("Hello world", "Spanish")          ' → Hola mundo
=AI_SENTIMENT("I love this product!")           ' → Positive
```

### 3. Process Entire Tables (Bulk Fill)

**Example table with clear headers:**
```
| Product Name    | Price   | Category    | Description          |
|----------------|---------|------------|---------------------|
| Widget A        | 29.99   |            | (leave empty)       |
| Widget B        | 45.50   |            | (leave empty)       |
```

**Steps:**
1. Select any cell in your table
2. Press `Ctrl+Shift+A` (hotkey)
3. Enter prompt: `"Fill in product categories and descriptions"`
4. Click **Run**

**Result:**
```
| Product Name    | Price   | Category    | Description          |
|----------------|---------|------------|---------------------|
| Widget A        | 29.99   | Electronics | Small electronic widget |
| Widget B        | 45.50   | Electronics | Medium electronic widget |
```

---

## Installation

### Quick Install

1. **Download** `OllamaLLM.xlam` from [Releases](../../releases)
2. **Save** to `%APPDATA%\Microsoft\AddIns\OllamaLLM.xlam`
3. **Enable** in Excel: `File → Options → Add-ins → Manage: Excel Add-ins → Go → Browse`

### First Run

On first use, the add-in creates a config file at:
```
%APPDATA%\OllamaLLM\config.ini
```

Edit this file to set your default model, endpoint, and API keys.

---

## Configuration

### Config File Location

```
%APPDATA%\OllamaLLM\config.ini
```

**Open via macro:** Press `Alt+F8` → Run `OllamaLLM.xlam!Open_AI_Config`

### Example Configuration

```ini
[ai]
model = qwen3:30b-a3b-instruct-2507-q8_0
endpoint = http://localhost:11434
api_key = 
temperature = 0.2
max_tokens = 512
system = You are a helpful assistant working inside Microsoft Excel. Return only the final answer with no extra words.

[search]
model = sonar-pro
endpoint = https://api.perplexity.ai
api_key = YOUR_PERPLEXITY_API_KEY
temperature = 0.2
max_tokens = 512
```

---

## Provider Setup

### Local (Ollama)

Free, private, runs on your machine.

```ini
[ai]
endpoint = http://localhost:11434
model = qwen3:30b-a3b-instruct-2507-q8_0
```

**Setup:**
```bash
# Install Ollama from https://ollama.ai
ollama pull qwen3:30b-a3b-instruct-2507-q8_0
ollama serve
```

### OpenAI

```ini
[ai]
endpoint = https://api.openai.com
model = gpt-4o-mini
api_key = sk-...
```

### Perplexity (with Web Search)

```ini
[search]
endpoint = https://api.perplexity.ai
model = sonar-pro
api_key = pplx-...
```

### Google Gemini (with Web Search)

```ini
[search]
endpoint = https://generativelanguage.googleapis.com/v1beta
model = gemini-2.0-flash
api_key = ...
```

### OpenRouter (Access Multiple Providers)

```ini
[ai]
endpoint = https://openrouter.ai/api
model = anthropic/claude-3.5-sonnet
api_key = sk-or-...
```

---

## Function Reference

### AI(prompt, [model], [temperature], [max_tokens], [endpoint], [api_key])

General-purpose AI function. Send any prompt, get a concise answer.

```excel
=AI("What is 15% of 847?")
=AI("Explain CAGR in one sentence", "llama3.1:8b")
```

### AI_SEARCH(prompt, ...)

Same as `AI()` but uses search-enabled models (Perplexity Sonar, Gemini with Google Search).

```excel
=AI_SEARCH("What is Apple's current stock price?")
=AI_SEARCH("Latest news about renewable energy")
```

### AI_EXTRACT(text, field, ...)

Extract a specific piece of information from text.

```excel
=AI_EXTRACT(A1, "email")           ' → john@example.com
=AI_EXTRACT(A1, "phone number")    ' → (555) 123-4567
=AI_EXTRACT(A1, "company name")    ' → Acme Corp
=AI_EXTRACT(A1, "price")           ' → $29.99
```

### AI_CLASSIFY(text, categories, ...)

Classify text into one of the provided categories.

```excel
' Comma-separated categories
=AI_CLASSIFY(A1, "Electronics, Clothing, Food, Other")

' Range of categories
=AI_CLASSIFY(A1, $D$1:$D$10)
```

### AI_TRANSLATE(text, targetLang, ...)

Translate text to any language.

```excel
=AI_TRANSLATE(A1, "Spanish")
=AI_TRANSLATE(A1, "Simplified Chinese")
=AI_TRANSLATE(A1, "French")
```

### AI_SUMMARIZE(text, [maxWords], ...)

Summarize text to a specified word count (default: 50 words).

```excel
=AI_SUMMARIZE(A1)       ' 50-word summary
=AI_SUMMARIZE(A1, 25)   ' 25-word summary
=AI_SUMMARIZE(A1, 100)  ' 100-word summary
```

### AI_SENTIMENT(text, ...)

Analyze the sentiment of text. Returns exactly one of: **Positive**, **Negative**, or **Neutral**.

```excel
=AI_SENTIMENT("I love this product!")        ' → Positive
=AI_SENTIMENT("This is the worst service")   ' → Negative
=AI_SENTIMENT("The package arrived today")   ' → Neutral
```

### AI_FIX(text, [rules], ...)

Fix grammar, spelling, and formatting issues.

```excel
=AI_FIX(A1)                           ' General fix
=AI_FIX(A1, "formal tone")            ' Apply formal tone
=AI_FIX(A1, "British English")        ' Use British spelling
=AI_FIX(A1, "remove passive voice")   ' Apply specific rule
```

---

## Requirements

- **Excel for Windows** (uses WinHTTP)
- **Ollama** (for local use) or API key for cloud providers
- For Ollama: Model must be pulled (`ollama pull <model>`)

---

## Build from Source

1. Open Excel → `Alt+F11` (VBA Editor)
2. Import all files from `/src`:
   - `modAI_Function.bas`
   - `modAI_Tooltips.bas`
   - `modAI_Bulk.bas`
   - `frmAIBulk.frm`
   - `JsonConverter.bas`
3. Enable **Microsoft Scripting Runtime** (Tools → References)
4. For `frmAIBulk`, add a checkbox named `chkPromptRow` with caption "Row 1 contains column-specific prompts"
5. Save as `.xlam` to `/add-in/OllamaLLM.xlam`

---

## Security

- **Local models (Ollama):** Your data never leaves your machine
- **Cloud providers:** Data is sent to the provider's API
- **No telemetry:** This add-in doesn't collect any usage data
- **Open source:** Audit the code yourself

If exposing Ollama beyond your LAN, protect it with a firewall or reverse proxy.

---

## Credits

- JSON parsing via [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) (MIT) by Tim Hall

---

## License

MIT — see [LICENSE](LICENSE)
