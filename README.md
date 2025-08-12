# Ollama Excel UDF — `AI()`

An Excel add-in (`.xlam`) that calls a local/remote Ollama server using the OpenAI-compatible `/v1/chat/completions` endpoint and returns short, cell-friendly answers.

---

## Formula

```excel
=AI(prompt, [model], [temperature], [max_tokens], [system], [endpoint])
```

---

## Quick Install

1. **Download the latest Release asset:**  
   [OllamaLLM.xlam](https://github.com/arsaboo/ollama-excel-udf/releases) (see Releases on this repo).

2. **In Excel:**  
   `File → Options → Add-ins → Manage: Excel Add-ins → Go… → Browse…`  
   Pick `OllamaLLM.xlam` and ensure it’s checked.

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
=AI("ping","qwen3:30b-a3b-instruct-2507-q8_0",0.2,128,"","http://192.168.2.50:11434")
```

*(Host-only is fine; `/v1/chat/completions` is auto-appended.)*

---

## Parameters

- **`prompt`** (required):  
  Your question/instruction (plain text).

- **`model`** (optional, default: `qwen3:30b-a3b-instruct-2507-q8_0`):  
  Must exist on the Ollama server (`ollama list`).

- **`temperature`** (optional, default: `0.2`):  
  `0.0–1.0`; lower = more deterministic (best for spreadsheets).

- **`max_tokens`** (optional, default: `512`):  
  Upper bound on response length.

- **`system`** (optional):  
  System prompt; default forces concise, single-value answers for Excel cells.

- **`endpoint`** (optional, default: `http://192.168.2.162:11434/v1/chat/completions`):  
  Full API URL or just `scheme://host:port`.

> **Note:**  
> Excel shows function help in the Function Arguments (`fx`) dialog, not inline while typing.

---

## Requirements

- **Excel for Windows** (uses WinHTTP).
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
