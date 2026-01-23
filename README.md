# Ollama Excel UDF — `AI()`

An Excel add-in (`.xlam`) that calls a local/remote Ollama server using the OpenAI-compatible `/v1/chat/completions` endpoint and returns short, cell-friendly answers.

---

## Formula

```excel
=AI(prompt, [model], [temperature], [max_tokens], [endpoint], [api_key])
=AI_SEARCH(prompt, [model], [temperature], [max_tokens], [endpoint], [api_key])
```

---

## Examples

### 1. Get Capitals

![Capitals Example](assets/Capitals.gif)

*Using:*
```excel
=AI("What is the capital of USA?")
```

### 2. Calculate Percentages

![Percentages Example](assets/Percentages.gif)

*Using:*
```excel
=AI("calculate 5% of "&A1)
```

---

## Quick Install

1. **Locate the add-in file:**
   Use the provided `OllamaLLM.xlam` file in the `/add-in` folder of this repository.

2. **In Excel:**
   `File → Options → Add-ins → Manage: Excel Add-ins → Go… → Browse…`
   Pick `OllamaLLM.xlam` from the `/add-in` folder and ensure it’s checked.

3. **(For Developers Only):**
   If building from source: enable `Tools → References → Microsoft Scripting Runtime` in the VBA editor.

---

## Usage

### Basic

```excel
=AI("What is the capital of USA?")
```

**Output:**
`Washington, D.C.`

---

### Change Model

```excel
=AI("Explain CAGR in one sentence","llama3.1:8b")
```

---

### Change Endpoint

```excel
=AI("ping","qwen3:30b-a3b-instruct-2507-q8_0",0.2,128,"http://192.168.2.50:11434")
```

*(Host-only is fine; `/v1/chat/completions` is auto-appended.)*

---

### Perplexity (API Key)

```excel
=AI("Hello!","sonar-pro",0.2,128,"https://api.perplexity.ai","YOUR_API_KEY")
```

### Search

```excel
=AI_SEARCH("latest GDP of Japan","sonar-pro",0.2,256,"https://api.perplexity.ai","YOUR_API_KEY")
```

---

## Provider Quick Reference

Use these values for the `model` and `endpoint` arguments (or set them in the AI Settings sheet).

- **OpenAI**
  - Endpoint: `https://api.openai.com`
  - Models (examples): `gpt-4o-mini`, `gpt-4.1-mini`
  - Notes: Requires API key.

- **Perplexity**
  - Endpoint: `https://api.perplexity.ai`
  - Models (examples): `sonar-pro`, `sonar-reasoning`
  - Notes: Requires API key.

- **OpenRouter**
  - Endpoint: `https://openrouter.ai/api`
  - Models (examples): `openai/gpt-4o-mini`, `anthropic/claude-3.5-sonnet`
  - Notes: Requires API key; some models may require additional headers.

---

## Parameters

- **`prompt`** (required):
  Your question/instruction (plain text).

- **`model`** (optional):
  Default from INI (`ai.model`), or `qwen3:30b-a3b-instruct-2507-q8_0` if missing.

- **`temperature`** (optional):
  Default from INI (`ai.temperature`), or `0.2` if missing.

- **`max_tokens`** (optional):
  Default from INI (`ai.max_tokens`), or `512` if missing.


- **`endpoint`** (optional):
  Default from INI (`ai.endpoint`), or built-in default if missing.

- **`api_key`** (optional):
  Default from INI (`ai.api_key`). Sent as `Authorization: Bearer <key>`.

---

## Defaults via INI

The add-in reads defaults from `%APPDATA%\OllamaLLM\config.ini` and auto-creates it if missing.

Example:

```ini
[ai]
model = qwen3:30b-a3b-instruct-2507-q8_0
endpoint = https://api.perplexity.ai
api_key = YOUR_API_KEY
temperature = 0.2
max_tokens = 512
system = You are a helpful assistant working inside Microsoft Excel. Return only the final answer with no extra words. Do not include explanations, context, or additional sentences. Use plain text only (no Markdown). If the answer is a single value, output only that value and its unit.

[search]
model = sonar-pro
endpoint = https://api.perplexity.ai
api_key = YOUR_API_KEY
temperature = 0.2
max_tokens = 512
system = You are a helpful assistant working inside Microsoft Excel. Return only the final answer with no extra words. Do not include explanations, context, or additional sentences. Use plain text only (no Markdown). If the answer is a single value, output only that value and its unit.
```

Gemini search example:

```ini
[search]
model = gemini-1.5-flash
endpoint = https://generativelanguage.googleapis.com/v1beta
api_key = YOUR_API_KEY
temperature = 0.2
max_tokens = 512
system = You are a helpful assistant working inside Microsoft Excel. Return only the final answer with no extra words. Do not include explanations, context, or additional sentences. Use plain text only (no Markdown). If the answer is a single value, output only that value and its unit.
```

OpenAI Responses API search example (for models like `gpt-5-mini`):

```ini
[search]
model = gpt-5-mini
endpoint = https://api.openai.com
api_key = YOUR_API_KEY
temperature = 0.2
max_tokens = 512
system = You are a helpful assistant working inside Microsoft Excel. Always return only the most concise, direct answer to the user's question. Do not include explanations, context, or extra words. Use plain text only (no Markdown). If the answer is a single value, output only that value.
```

To open the config file via macro:

- Press `Alt+F8`
- In **Macro name**, type `OllamaLLM.xlam!Open_AI_Config`
- Click **Run**

> **Note:**
> Excel shows function help in the Function Arguments (`fx`) dialog, not inline while typing.

---

## Requirements

- **Excel for Windows** (uses WinHTTP).
  *This add-in has only been tested on Windows.*
- Reachable **Ollama server** (default: `http://192.168.2.162:11434`).
  If remote, start server with `OLLAMA_HOST=0.0.0.0` and open TCP 11434.
- **Model pulled:**
  ```sh
  ollama pull qwen3:30b-a3b-instruct-2507-q8_0
  ```

---

## Build from Source

1. In Excel (`Alt+F11`), import files under `/src`:
    - `modAI_Functions.bas`
    - `modAI_Tooltips.bas`
    - `JsonConverter.bas` (from [VBA-JSON](https://github.com/VBA-tools/VBA-JSON))
2. Enable **Microsoft Scripting Runtime** in VBA editor.
3. Save as `.xlam` under `/add-in/OllamaLLM.xlam`.
4. Re-open Excel. The add-in auto-registers UDF tooltips.

---

## Security Notes

- This add-in talks directly to your Ollama host; **no API key is used**.
- If exposing beyond LAN, **protect the host** (firewall, reverse proxy, auth).
- Consider **signing the add-in** with `SelfCert.exe` to reduce macro warnings.

---

## Credits

- JSON parsing via [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) (MIT) by Tim Hall.

---

## License

MIT — see [LICENSE](LICENSE).
