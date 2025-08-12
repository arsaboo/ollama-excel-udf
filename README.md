# Ollama Excel UDF — `AI()`

An Excel add-in (.xlam) that calls a local/remote **Ollama** server using the OpenAI-compatible
`/v1/chat/completions` endpoint and returns short, cell-friendly answers.

<p align="center"><b>Formula</b>: `=AI(prompt, [model], [temperature], [max_tokens], [system], [endpoint])`</p>

---

## Quick install

1. Download the latest **Release** asset: `OllamaLLM.xlam` (see Releases on this repo).
2. In Excel: **File → Options → Add-ins → Manage: Excel Add-ins → Go… → Browse…**  
   Pick `OllamaLLM.xlam` and ensure it’s checked.
3. (Dev only) If building from source: enable **Tools → References → Microsoft Scripting Runtime**.

## Usage

**Basic**
```excel
=AI("What is the capital of USA?")
