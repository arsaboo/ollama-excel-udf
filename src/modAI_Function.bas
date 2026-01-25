Attribute VB_Name = "modAI_Function"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long
#Else
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long
#End If

' === AI worksheet function ===
' NOTE: Requires Tools ? References ? Microsoft Scripting Runtime
Public Function AI(prompt As String, _
                   Optional model As String = "", _
                   Optional temperature As Variant, _
                   Optional max_tokens As Variant, _
                   Optional endpoint As String = "", _
                   Optional api_key As String = "") As String
    Dim system As String
    Dim resolvedModel As String
    Dim resolvedEndpoint As String
    Dim resolvedApiKey As String
    Dim resolvedTemp As Double
    Dim resolvedTokens As Long

    ' Initialize provider defaults if needed
    InitializeProviderDefaults

    ' Use provider settings from modProvider if not explicitly provided
    If Len(model) > 0 Then
        resolvedModel = model
    Else
        resolvedModel = GetCurrentModel()
    End If

    If Len(endpoint) > 0 Then
        resolvedEndpoint = endpoint
    Else
        resolvedEndpoint = GetCurrentEndpoint()
    End If

    If Len(api_key) > 0 Then
        resolvedApiKey = api_key
    Else
        resolvedApiKey = GetCurrentApiKey()
    End If

    If IsMissing(temperature) Or IsEmpty(temperature) Then
        resolvedTemp = GetCurrentTemperature()
    Else
        resolvedTemp = CDbl(temperature)
    End If

    If IsMissing(max_tokens) Or IsEmpty(max_tokens) Then
        resolvedTokens = GetCurrentMaxTokens()
    Else
        resolvedTokens = CLng(max_tokens)
    End If

    system = GetCurrentSystem()

    AI = AI_Core(prompt, system, resolvedModel, resolvedTemp, resolvedTokens, resolvedEndpoint, resolvedApiKey)
End Function

' === Helper to resolve provider parameters ===
Private Sub ResolveProviderParams(ByVal model As String, _
                                  ByVal temperature As Variant, _
                                  ByVal max_tokens As Variant, _
                                  ByVal endpoint As String, _
                                  ByVal api_key As String, _
                                  ByRef resolvedModel As String, _
                                  ByRef resolvedTemp As Double, _
                                  ByRef resolvedTokens As Long, _
                                  ByRef resolvedEndpoint As String, _
                                  ByRef resolvedApiKey As String)

    If Len(model) > 0 Then
        resolvedModel = model
    Else
        resolvedModel = GetCurrentModel()
    End If

    If Len(endpoint) > 0 Then
        resolvedEndpoint = endpoint
    Else
        resolvedEndpoint = GetCurrentEndpoint()
    End If

    If Len(api_key) > 0 Then
        resolvedApiKey = api_key
    Else
        resolvedApiKey = GetCurrentApiKey()
    End If

    If IsMissing(temperature) Or IsEmpty(temperature) Then
        resolvedTemp = GetCurrentTemperature()
    Else
        resolvedTemp = CDbl(temperature)
    End If

    If IsMissing(max_tokens) Or IsEmpty(max_tokens) Then
        resolvedTokens = GetCurrentMaxTokens()
    Else
        resolvedTokens = CLng(max_tokens)
    End If
End Sub

' === Core AI function (internal) ===
' All AI_* functions call this shared implementation
Private Function AI_Core(prompt As String, _
                         systemPrompt As String, _
                         model As String, _
                         temperature As Double, _
                         max_tokens As Long, _
                         endpoint As String, _
                         api_key As String) As String
    Dim http As Object
    Dim status As Long
    Dim body As String
    Dim payload As String
    Dim json As Object
    Dim content As String
    Dim url As String
    Dim startTime As Double
    Dim durationMs As Long
    Dim finishReason As String
    Dim resultStatus As String

    url = NormalizeEndpoint(endpoint)

    On Error GoTo FailSoft

    Dim thinkEnabled As Boolean
    thinkEnabled = GetCurrentThink()
    If IsOllamaEndpoint(endpoint) Or IsGptOssModel(model) Then
    payload = BuildChatPayload(prompt, model, temperature, max_tokens, systemPrompt, thinkEnabled)
    Else
        payload = BuildChatPayload(prompt, model, temperature, max_tokens, systemPrompt)
    End If

    startTime = Timer

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 30000, 30000, 30000, 120000
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Accept", "application/json"
    If Len(api_key) > 0 Then
        http.SetRequestHeader "Authorization", "Bearer " & api_key
    End If
    http.Send payload

    status = http.status
    body = http.responseText
    durationMs = CLng((Timer - startTime) * 1000)

    If status <> 200 Then
        resultStatus = "Error: HTTP " & status & " - " & http.StatusText
        AI_Core = resultStatus & ". Body: " & Left$(body, 500)
        LogAIResponse "AI_Core", model, url, payload, durationMs, status, body, 0, 0, "", resultStatus
        Exit Function
    End If

    Set json = JsonConverter.ParseJson(body)
    On Error Resume Next
    content = json("choices")(1)("message")("content")
    Dim thinkingContent As String
    thinkingContent = ""
    thinkingContent = json("choices")(1)("message")("thinking")
    If Len(thinkingContent) = 0 Then
        thinkingContent = json("choices")(1)("message")("reasoning")
    End If
    finishReason = ""
    finishReason = json("choices")(1)("finish_reason")
    On Error GoTo FailSoft

    If Len(content) = 0 Then
        If LCase$(finishReason) = "length" Then
            resultStatus = "Error: Token limit reached. " & _
                           "Increase max_tokens."
            AI_Core = resultStatus
        ElseIf Len(thinkingContent) > 0 Then
            resultStatus = "Error: Model returned thinking trace but no final answer. " & _
                           "Increase max_tokens or wait for model to complete."
            AI_Core = resultStatus
        Else
            resultStatus = "Error: Missing content in response."
            AI_Core = resultStatus & " Raw: " & Left$(body, 500)
        End If
    Else
        resultStatus = "OK"
        AI_Core = Trim(content)
    End If

    LogAIResponse "AI_Core", model, url, payload, durationMs, status, body, Len(content), Len(thinkingContent), finishReason, resultStatus
    Exit Function

FailSoft:
    resultStatus = "VBA Error #" & Err.Number & ": " & Err.Description
    AI_Core = resultStatus
    LogAIResponse "AI_Core", model, url, payload, durationMs, status, body, Len(content), 0, finishReason, resultStatus
End Function

Public Function AI_SEARCH(prompt As String, _
                          Optional model As String = "", _
                          Optional temperature As Variant, _
                          Optional max_tokens As Variant, _
                          Optional endpoint As String = "", _
                          Optional api_key As String = "") As String
    Dim http As Object
    Dim status As Long
    Dim body As String
    Dim payload As String
    Dim json As Object
    Dim content As String
    Dim url As String
    Dim isGemini As Boolean
    Dim isResponses As Boolean
    Dim system As String
    Dim resolvedTemp As Double
    Dim resolvedTokens As Long
    Dim startTime As Double
    Dim durationMs As Long
    Dim finishReason As String
    Dim resultStatus As String
    Dim provider As String

    ' Initialize provider defaults if needed
    InitializeProviderDefaults

    ' Use search provider settings from modProvider if not explicitly provided
    If Len(model) = 0 Then model = GetSearchModel()
    If Len(endpoint) = 0 Then endpoint = GetSearchEndpoint()
    If Len(api_key) = 0 Then api_key = GetSearchApiKey()

    isGemini = InStr(1, endpoint, "generativelanguage.googleapis.com", vbTextCompare) > 0
    isResponses = IsResponsesModel(model, endpoint)

    If IsMissing(temperature) Or IsEmpty(temperature) Then
        resolvedTemp = GetSearchTemperature()
    Else
        resolvedTemp = CDbl(temperature)
    End If

    If IsMissing(max_tokens) Or IsEmpty(max_tokens) Then
        resolvedTokens = GetSearchMaxTokens()
    Else
        resolvedTokens = CLng(max_tokens)
    End If

    system = GetSearchSystem()

    If isGemini Then
        url = NormalizeGeminiEndpoint(endpoint, model, api_key)
        provider = "gemini"
    ElseIf isResponses Then
        url = NormalizeResponsesEndpoint(endpoint)
        provider = "responses"
    Else
        url = NormalizeEndpoint(endpoint)
        provider = "chat"
    End If

    On Error GoTo FailSoft

    If isGemini Then
        payload = BuildGeminiPayload(prompt, system)
    ElseIf isResponses Then
        payload = BuildResponsesPayload(prompt, model, resolvedTemp, resolvedTokens, system)
    Else
        Dim thinkEnabled As Boolean
        thinkEnabled = GetCurrentThink()
        If IsOllamaEndpoint(endpoint) Then
            payload = BuildChatPayload(prompt, model, resolvedTemp, resolvedTokens, system, thinkEnabled)
        Else
            payload = BuildChatPayload(prompt, model, resolvedTemp, resolvedTokens, system)
        End If
    End If

    startTime = Timer

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 30000, 30000, 30000, 120000
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Accept", "application/json"
    If Len(api_key) > 0 And Not isGemini Then
        http.SetRequestHeader "Authorization", "Bearer " & api_key
    End If
    http.Send payload

    status = http.status
    body = http.responseText
    durationMs = CLng((Timer - startTime) * 1000)

    If status <> 200 Then
        resultStatus = "Error: HTTP " & status & " - " & http.StatusText
        AI_SEARCH = resultStatus & ". Body: " & Left$(body, 500)
        LogAIResponse "AI_SEARCH(" & provider & ")", model, url, payload, durationMs, status, body, 0, 0, "", resultStatus
        Exit Function
    End If

    Set json = JsonConverter.ParseJson(body)
    On Error Resume Next
    If isGemini Then
        content = json("candidates")(1)("content")("parts")(1)("text")
    ElseIf isResponses Then
        On Error Resume Next
        content = json("output")(1)("content")(1)("text")
        If Len(content) = 0 Then
            content = json("output_text")
        End If
        On Error GoTo FailSoft
    Else
        content = json("choices")(1)("message")("content")
    End If
    finishReason = ""
    If Not isGemini And Not isResponses Then
        finishReason = json("choices")(1)("finish_reason")
    End If
    On Error GoTo FailSoft

    If Len(content) = 0 Then
        resultStatus = "Error: Missing content in response."
        AI_SEARCH = resultStatus & " Raw: " & Left$(body, 500)
    Else
        resultStatus = "OK"
        AI_SEARCH = Trim(content)
    End If
    LogAIResponse "AI_SEARCH(" & provider & ")", model, url, payload, durationMs, status, body, Len(content), 0, finishReason, resultStatus
    Exit Function

FailSoft:
    resultStatus = "VBA Error #" & Err.Number & ": " & Err.Description
    AI_SEARCH = resultStatus
    LogAIResponse "AI_SEARCH(" & provider & ")", model, url, payload, durationMs, status, body, Len(content), 0, finishReason, resultStatus
End Function

Public Function AI_Version() As String
    AI_Version = "2026-01-25.1"
End Function

' === AI_EXTRACT: Extract a specific field from text ===
Public Function AI_EXTRACT(text As String, _
                           field As String, _
                           Optional model As String = "", _
                           Optional temperature As Variant, _
                           Optional max_tokens As Variant, _
                           Optional endpoint As String = "", _
                           Optional api_key As String = "") As String
    Dim systemPrompt As String
    Dim resolvedModel As String
    Dim resolvedEndpoint As String
    Dim resolvedApiKey As String
    Dim resolvedTemp As Double
    Dim resolvedTokens As Long

    InitializeProviderDefaults
    ResolveProviderParams model, temperature, max_tokens, endpoint, api_key, _
                          resolvedModel, resolvedTemp, resolvedTokens, resolvedEndpoint, resolvedApiKey

    systemPrompt = "Extract only the " & field & " from the following text. " & _
                   "Return just the extracted value with no additional text. " & _
                   "If the field cannot be found, return empty."

    AI_EXTRACT = AI_Core(text, systemPrompt, resolvedModel, resolvedTemp, resolvedTokens, resolvedEndpoint, resolvedApiKey)
End Function

' === AI_CLASSIFY: Classify text into one of the provided categories ===
Public Function AI_CLASSIFY(text As String, _
                            categories As Variant, _
                            Optional model As String = "", _
                            Optional temperature As Variant, _
                            Optional max_tokens As Variant, _
                            Optional endpoint As String = "", _
                            Optional api_key As String = "") As String
    Dim systemPrompt As String
    Dim categoryList As String
    Dim resolvedModel As String
    Dim resolvedEndpoint As String
    Dim resolvedApiKey As String
    Dim resolvedTemp As Double
    Dim resolvedTokens As Long

    InitializeProviderDefaults
    ResolveProviderParams model, temperature, max_tokens, endpoint, api_key, _
                          resolvedModel, resolvedTemp, resolvedTokens, resolvedEndpoint, resolvedApiKey

    categoryList = ParseCategories(categories)

    systemPrompt = "Classify the following text into exactly one of these categories: " & categoryList & ". " & _
                   "Return only the category name, nothing else."

    AI_CLASSIFY = AI_Core(text, systemPrompt, resolvedModel, resolvedTemp, resolvedTokens, resolvedEndpoint, resolvedApiKey)
End Function

' === AI_TRANSLATE: Translate text to target language ===
Public Function AI_TRANSLATE(text As String, _
                             targetLang As String, _
                             Optional model As String = "", _
                             Optional temperature As Variant, _
                             Optional max_tokens As Variant, _
                             Optional endpoint As String = "", _
                             Optional api_key As String = "") As String
    Dim systemPrompt As String
    Dim resolvedModel As String
    Dim resolvedEndpoint As String
    Dim resolvedApiKey As String
    Dim resolvedTemp As Double
    Dim resolvedTokens As Long

    InitializeProviderDefaults
    ResolveProviderParams model, temperature, max_tokens, endpoint, api_key, _
                          resolvedModel, resolvedTemp, resolvedTokens, resolvedEndpoint, resolvedApiKey

    systemPrompt = "Translate the following text to " & targetLang & ". " & _
                   "Return only the translation with no explanations or notes."

    AI_TRANSLATE = AI_Core(text, systemPrompt, resolvedModel, resolvedTemp, resolvedTokens, resolvedEndpoint, resolvedApiKey)
End Function

' === AI_SUMMARIZE: Summarize text to specified word count ===
Public Function AI_SUMMARIZE(text As String, _
                             Optional maxWords As Long = 50, _
                             Optional model As String = "", _
                             Optional temperature As Variant, _
                             Optional max_tokens As Variant, _
                             Optional endpoint As String = "", _
                             Optional api_key As String = "") As String
    Dim systemPrompt As String
    Dim resolvedModel As String
    Dim resolvedEndpoint As String
    Dim resolvedApiKey As String
    Dim resolvedTemp As Double
    Dim resolvedTokens As Long

    InitializeProviderDefaults
    ResolveProviderParams model, temperature, max_tokens, endpoint, api_key, _
                          resolvedModel, resolvedTemp, resolvedTokens, resolvedEndpoint, resolvedApiKey

    systemPrompt = "Summarize the following text in " & maxWords & " words or fewer. " & _
                   "Return only the summary, no preamble or additional text."

    AI_SUMMARIZE = AI_Core(text, systemPrompt, resolvedModel, resolvedTemp, resolvedTokens, resolvedEndpoint, resolvedApiKey)
End Function

' === AI_SENTIMENT: Analyze sentiment of text ===
Public Function AI_SENTIMENT(text As String, _
                             Optional model As String = "", _
                             Optional temperature As Variant, _
                             Optional max_tokens As Variant, _
                             Optional endpoint As String = "", _
                             Optional api_key As String = "") As String
    Dim systemPrompt As String
    Dim resolvedModel As String
    Dim resolvedEndpoint As String
    Dim resolvedApiKey As String
    Dim resolvedTemp As Double
    Dim resolvedTokens As Long

    InitializeProviderDefaults
    ResolveProviderParams model, temperature, max_tokens, endpoint, api_key, _
                          resolvedModel, resolvedTemp, resolvedTokens, resolvedEndpoint, resolvedApiKey

    systemPrompt = "Analyze the sentiment of the following text. " & _
                   "Return exactly one word: Positive, Negative, or Neutral."

    AI_SENTIMENT = AI_Core(text, systemPrompt, resolvedModel, resolvedTemp, resolvedTokens, resolvedEndpoint, resolvedApiKey)
End Function

' === AI_FIX: Fix grammar, spelling, and formatting ===
Public Function AI_FIX(text As String, _
                       Optional rules As String = "", _
                       Optional model As String = "", _
                       Optional temperature As Variant, _
                       Optional max_tokens As Variant, _
                       Optional endpoint As String = "", _
                       Optional api_key As String = "") As String
    Dim systemPrompt As String
    Dim resolvedModel As String
    Dim resolvedEndpoint As String
    Dim resolvedApiKey As String
    Dim resolvedTemp As Double
    Dim resolvedTokens As Long

    InitializeProviderDefaults
    ResolveProviderParams model, temperature, max_tokens, endpoint, api_key, _
                          resolvedModel, resolvedTemp, resolvedTokens, resolvedEndpoint, resolvedApiKey

    systemPrompt = "Fix any grammar, spelling, and formatting issues in the following text. "
    If Len(rules) > 0 Then
        systemPrompt = systemPrompt & "Apply these additional rules: " & rules & ". "
    End If
    systemPrompt = systemPrompt & "Return only the corrected text."

    AI_FIX = AI_Core(text, systemPrompt, resolvedModel, resolvedTemp, resolvedTokens, resolvedEndpoint, resolvedApiKey)
End Function

' === Helper: Parse categories from string or range ===
Private Function ParseCategories(categories As Variant) As String
    Dim result As String
    Dim cell As Range
    Dim cellValue As String

    result = ""

    If TypeName(categories) = "Range" Then
        ' Range: iterate cells and join with commas
        For Each cell In categories
            cellValue = Trim$(CStr(cell.Value))
            If Len(cellValue) > 0 Then
                If Len(result) > 0 Then result = result & ", "
                result = result & cellValue
            End If
        Next cell
    Else
        ' String: use as-is
        result = CStr(categories)
    End If

    ParseCategories = result
End Function

' Build OpenAI-compatible payload (uses strongly-typed Dictionary for VBA-JSON)
Private Function BuildChatPayload(prompt As String, _
                                  model As String, _
                                  temperature As Double, _
                                  max_tokens As Long, _
                                  system As String, _
                                  Optional think As Variant) As String
    Dim root As Scripting.Dictionary
    Dim messages As Collection
    Dim msg As Scripting.Dictionary

    Set root = New Scripting.Dictionary
    Set messages = New Collection

    If Len(system) > 0 Then
        Set msg = New Scripting.Dictionary
        msg.Add "role", "system"
        msg.Add "content", system
        messages.Add msg
    End If

    Set msg = New Scripting.Dictionary
    msg.Add "role", "user"
    msg.Add "content", prompt
    messages.Add msg

    root.Add "model", model
    root.Add "messages", messages
    root.Add "temperature", temperature
    root.Add "max_tokens", max_tokens
    root.Add "stream", False
    If IsGptOssModel(model) Then
        root.Add "think", "low"
    ElseIf Not IsMissing(think) Then
        root.Add "think", CBool(think)
    End If

    BuildChatPayload = JsonConverter.ConvertToJson(root, Whitespace:=0)
End Function

' Accepts host-only or full path; appends /chat/completions or /v1/chat/completions if needed
Private Function NormalizeEndpoint(ByVal e As String) As String
    Dim s As String
    s = Trim(e)
    If Len(s) = 0 Then
        s = "http://127.0.0.1:11434/v1/chat/completions"
    End If
    ' If it ends with /api/chat, /v1/chat/completions, or /chat/completions, leave as is
    If Right$(s, 14) = "/api/chat" Or Right$(s, 21) = "/v1/chat/completions" Or Right$(s, 18) = "/chat/completions" Then
        NormalizeEndpoint = s
        Exit Function
    End If
    ' If it looks like just scheme://host[:port] or with trailing slash, append path
    If InStr(1, s, "/v1/chat/completions", vbTextCompare) = 0 And _
       InStr(1, s, "/api/chat", vbTextCompare) = 0 And _
       InStr(1, s, "/chat/completions", vbTextCompare) = 0 Then
        If Right$(s, 1) = "/" Then
            s = Left$(s, Len(s) - 1)
        End If
        If InStr(1, s, "perplexity.ai", vbTextCompare) > 0 Then
            s = s & "/chat/completions"
        Else
            s = s & "/v1/chat/completions"
        End If
    End If
    NormalizeEndpoint = s
End Function

Private Function ResolveIniString(ByVal value As String, ByVal keyName As String, ByVal fallback As String, Optional ByVal sectionName As String = "ai") As String
    Dim settingValue As String
    settingValue = ReadIniValue(sectionName, keyName, fallback)
    If Len(value) > 0 Then
        ResolveIniString = value
    Else
        ResolveIniString = settingValue
    End If
End Function

Private Function ResolveIniDouble(ByVal keyName As String, ByVal fallback As Double, Optional ByVal sectionName As String = "ai") As Double
    Dim settingValue As String
    settingValue = ReadIniValue(sectionName, keyName, CStr(fallback))
    If Len(settingValue) > 0 Then
        ResolveIniDouble = CDbl(settingValue)
    Else
        ResolveIniDouble = fallback
    End If
End Function

Private Function ResolveIniLong(ByVal keyName As String, ByVal fallback As Long, Optional ByVal sectionName As String = "ai") As Long
    Dim settingValue As String
    settingValue = ReadIniValue(sectionName, keyName, CStr(fallback))
    If Len(settingValue) > 0 Then
        ResolveIniLong = CLng(settingValue)
    Else
        ResolveIniLong = fallback
    End If
End Function

Private Function GetCurrentThink() As Boolean
    Dim settingValue As String
    settingValue = LCase$(ReadIniValue("ai", "think", "false"))
    GetCurrentThink = (settingValue = "true" Or settingValue = "1" Or settingValue = "yes")
End Function

Private Function ReadIniValue(ByVal sectionName As String, ByVal keyName As String, ByVal fallback As String) As String
    Dim buffer As String
    Dim length As Long
    buffer = String$(1024, vbNullChar)
    length = GetPrivateProfileString(sectionName, keyName, fallback, buffer, Len(buffer), GetIniPath())
    ReadIniValue = Left$(buffer, length)
End Function

Private Sub EnsureIniDefaults()
    Dim iniPath As String
    Dim folderPath As String

    iniPath = GetIniPath()
    folderPath = Left$(iniPath, InStrRev(iniPath, "\") - 1)
    EnsureFolderExists folderPath

    WriteIniDefault "ai", "model", "qwen3:30b-a3b-instruct-2507-q8_0"
    WriteIniDefault "ai", "endpoint", "http://192.168.2.162:11434/v1/chat/completions"
    WriteIniDefault "ai", "api_key", ""
    WriteIniDefault "ai", "temperature", "0.2"
    WriteIniDefault "ai", "max_tokens", "512"
    WriteIniDefault "ai", "think", "false"
    WriteIniDefault "ai", "debug", "false"
    WriteIniDefault "ai", "system", DefaultSystemPrompt()
    WriteIniDefault "ai", "initialized", "0"

    WriteIniDefault "search", "model", "sonar-pro"
    WriteIniDefault "search", "endpoint", "https://api.perplexity.ai"
    WriteIniDefault "search", "api_key", ""
    WriteIniDefault "search", "temperature", "0.2"
    WriteIniDefault "search", "max_tokens", "512"
    WriteIniDefault "search", "system", DefaultSystemPrompt()
End Sub

Private Function DefaultSystemPrompt() As String
    DefaultSystemPrompt = "You are a helpful assistant working inside Microsoft Excel. " & _
                          "Return only the final answer with no extra words. " & _
                          "Do not include explanations, context, or additional sentences. " & _
                          "Use plain text only (no Markdown). " & _
                          "If the answer is a single value, output only that value and its unit."
End Function

Private Sub WriteIniDefault(ByVal sectionName As String, ByVal keyName As String, ByVal defaultValue As String)
    Dim existing As String
    existing = ReadIniValue(sectionName, keyName, "")
    If Len(existing) = 0 Then
        WritePrivateProfileString sectionName, keyName, defaultValue, GetIniPath()
    End If
End Sub

Private Function GetIniPath() As String
    GetIniPath = Environ$("APPDATA") & "\OllamaLLM\config.ini"
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

Private Function NormalizeGeminiEndpoint(ByVal e As String, ByVal model As String, ByVal apiKey As String) As String
    Dim s As String
    Dim base As String

    s = Trim(e)
    If Len(s) = 0 Then
        s = "https://generativelanguage.googleapis.com/v1beta"
    End If

    If InStr(1, s, ":generateContent", vbTextCompare) > 0 Then
        NormalizeGeminiEndpoint = AppendGeminiKey(s, apiKey)
        Exit Function
    End If

    base = s
    If Right$(base, 1) = "/" Then
        base = Left$(base, Len(base) - 1)
    End If

    NormalizeGeminiEndpoint = AppendGeminiKey(base & "/models/" & model & ":generateContent", apiKey)
End Function

Private Function NormalizeResponsesEndpoint(ByVal e As String) As String
    Dim s As String
    s = Trim(e)
    If Len(s) = 0 Then
        s = "https://api.openai.com/v1/responses"
    End If
    If InStr(1, s, "/v1/responses", vbTextCompare) > 0 Then
        NormalizeResponsesEndpoint = s
        Exit Function
    End If
    If Right$(s, 1) = "/" Then
        s = Left$(s, Len(s) - 1)
    End If
    NormalizeResponsesEndpoint = s & "/v1/responses"
End Function

Private Function IsResponsesModel(ByVal model As String, ByVal endpoint As String) As Boolean
    Dim m As String
    Dim e As String

    m = LCase$(Trim$(model))
    e = LCase$(Trim$(endpoint))

    If InStr(1, e, "api.openai.com", vbTextCompare) = 0 Then
        IsResponsesModel = False
        Exit Function
    End If

    If Left$(m, 5) = "gpt-5" Then
        IsResponsesModel = True
    Else
        IsResponsesModel = False
    End If
End Function

Private Function GetDebugEnabled() As Boolean
    Dim settingValue As String
    settingValue = LCase$(ReadIniValue("ai", "debug", "false"))
    GetDebugEnabled = (settingValue = "true" Or settingValue = "1" Or settingValue = "yes")
End Function

Private Function GetDebugLogPath() As String
    GetDebugLogPath = Environ$("APPDATA") & "\OllamaLLM\debug.log"
End Function

Private Function LogTimestamp() As String
    LogTimestamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function

Private Sub RotateLogIfNeeded()
    Dim logPath As String
    Dim fso As Object
    Dim logFile As Object
    Dim lastModified As Date

    logPath = GetDebugLogPath()
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FileExists(logPath) Then
        Set logFile = fso.GetFile(logPath)
        lastModified = logFile.DateLastModified
        If DateValue(lastModified) < DateValue(Now) Then
            fso.DeleteFile logPath
        End If
    End If
End Sub

Private Sub LogDebug(ByVal message As String, ByVal forceLog As Boolean)
    Dim logPath As String
    Dim fso As Object
    Dim ts As Object

    If Not forceLog And Not GetDebugEnabled() Then Exit Sub

    On Error Resume Next
    RotateLogIfNeeded
    logPath = GetDebugLogPath()
    Set fso = CreateObject("Scripting.FileSystemObject")
    EnsureFolderExists Left$(logPath, InStrRev(logPath, "\") - 1)
    Set ts = fso.OpenTextFile(logPath, 8, True)
    ts.WriteLine message
    ts.Close
    On Error GoTo 0
End Sub

Private Sub LogAIResponse(ByVal funcName As String, _
                          ByVal model As String, _
                          ByVal endpoint As String, _
                          ByVal payload As String, _
                          ByVal durationMs As Long, _
                          ByVal httpStatus As Long, _
                          ByVal responseBody As String, _
                          ByVal contentLength As Long, _
                          ByVal thinkingLength As Long, _
                          ByVal finishReason As String, _
                          ByVal resultStatus As String)
    Dim logMsg As String
    Dim truncatedResponse As String
    Dim forceLog As Boolean

    forceLog = (Left$(resultStatus, 5) = "Error" Or Left$(resultStatus, 9) = "VBA Error")
    If Not forceLog And Not GetDebugEnabled() Then Exit Sub

    If Len(responseBody) > 2000 Then
        truncatedResponse = Left$(responseBody, 2000) & "... [TRUNCATED]"
    Else
        truncatedResponse = responseBody
    End If

    logMsg = String$(80, "=") & vbCrLf & _
             "[" & LogTimestamp() & "] " & funcName & vbCrLf & _
             String$(80, "=") & vbCrLf & _
             "MODEL: " & model & vbCrLf & _
             "ENDPOINT: " & endpoint & vbCrLf & _
             "PAYLOAD:" & vbCrLf & _
             payload & vbCrLf & _
             "--- RESPONSE ---" & vbCrLf & _
             "DURATION: " & durationMs & " ms" & vbCrLf & _
             "HTTP_STATUS: " & httpStatus & vbCrLf & _
             "RESPONSE_SIZE: " & Len(responseBody) & " bytes" & vbCrLf & _
             "FINISH_REASON: " & IIf(Len(finishReason) > 0, finishReason, "MISSING") & vbCrLf & _
             "CONTENT_LENGTH: " & contentLength & " chars" & vbCrLf & _
             "THINKING_LENGTH: " & thinkingLength & " chars" & vbCrLf & _
             "RESULT: " & resultStatus & vbCrLf & _
             "[RAW RESPONSE]" & vbCrLf & _
             truncatedResponse

    LogDebug logMsg, forceLog
End Sub

Private Function IsGptOssModel(ByVal model As String) As Boolean
    IsGptOssModel = (InStr(1, LCase$(Trim$(model)), "gpt-oss", vbTextCompare) > 0)
End Function

Private Function IsOllamaEndpoint(ByVal endpoint As String) As Boolean
    Dim e As String
    e = LCase$(Trim$(endpoint))
    If InStr(1, e, "11434", vbTextCompare) > 0 Then
        IsOllamaEndpoint = True
    ElseIf InStr(1, e, "ollama", vbTextCompare) > 0 Then
        IsOllamaEndpoint = True
    ElseIf InStr(1, e, "/api/chat", vbTextCompare) > 0 Then
        IsOllamaEndpoint = True
    Else
        IsOllamaEndpoint = False
    End If
End Function

Private Function AppendGeminiKey(ByVal url As String, ByVal apiKey As String) As String
    If Len(apiKey) = 0 Then
        AppendGeminiKey = url
    ElseIf InStr(1, url, "?", vbTextCompare) > 0 Then
        AppendGeminiKey = url & "&key=" & apiKey
    Else
        AppendGeminiKey = url & "?key=" & apiKey
    End If
End Function

Private Function BuildGeminiPayload(ByVal prompt As String, ByVal system As String) As String
    Dim root As Scripting.Dictionary
    Dim contents As Collection
    Dim parts As Collection
    Dim contentObj As Scripting.Dictionary
    Dim partObj As Scripting.Dictionary
    Dim systemObj As Scripting.Dictionary
    Dim systemParts As Collection
    Dim systemPart As Scripting.Dictionary
    Dim tools As Collection
    Dim toolObj As Scripting.Dictionary
    Dim searchTool As Scripting.Dictionary

    Set root = New Scripting.Dictionary
    Set contents = New Collection
    Set parts = New Collection

    Set partObj = New Scripting.Dictionary
    partObj.Add "text", prompt
    parts.Add partObj

    Set contentObj = New Scripting.Dictionary
    contentObj.Add "role", "user"
    contentObj.Add "parts", parts
    contents.Add contentObj

    root.Add "contents", contents

    If Len(system) > 0 Then
        Set systemObj = New Scripting.Dictionary
        Set systemParts = New Collection
        Set systemPart = New Scripting.Dictionary
        systemPart.Add "text", system
        systemParts.Add systemPart
        systemObj.Add "parts", systemParts
        root.Add "systemInstruction", systemObj
    End If

    Set tools = New Collection
    Set toolObj = New Scripting.Dictionary
    Set searchTool = New Scripting.Dictionary
    toolObj.Add "google_search", searchTool
    tools.Add toolObj
    root.Add "tools", tools

    BuildGeminiPayload = JsonConverter.ConvertToJson(root, Whitespace:=0)
End Function

Private Function BuildResponsesPayload(ByVal prompt As String, ByVal model As String, ByVal temperature As Double, ByVal max_tokens As Long, ByVal system As String) As String
    Dim root As Scripting.Dictionary
    Dim inputItems As Collection
    Dim messageObj As Scripting.Dictionary
    Dim contentItems As Collection
    Dim contentObj As Scripting.Dictionary

    Set root = New Scripting.Dictionary
    Set inputItems = New Collection
    Set messageObj = New Scripting.Dictionary
    Set contentItems = New Collection
    Set contentObj = New Scripting.Dictionary

    contentObj.Add "type", "input_text"
    contentObj.Add "text", prompt
    contentItems.Add contentObj

    messageObj.Add "role", "user"
    messageObj.Add "content", contentItems
    inputItems.Add messageObj

    root.Add "model", model
    root.Add "input", inputItems
    root.Add "temperature", temperature
    root.Add "max_output_tokens", max_tokens

    If Len(system) > 0 Then
        root.Add "instructions", system
    End If

    BuildResponsesPayload = JsonConverter.ConvertToJson(root, Whitespace:=0)
End Function

Public Sub Open_AI_Config()
    Dim iniPath As String
    EnsureIniDefaults
    iniPath = GetIniPath()
    Shell "notepad.exe """ & iniPath & """", vbNormalFocus
End Sub

Public Sub AI_Notify_First_Run()
    Dim initialized As String
    Dim iniPath As String

    EnsureIniDefaults
    initialized = ReadIniValue("ai", "initialized", "0")
    If initialized <> "1" Then
        WritePrivateProfileString "ai", "initialized", "1", GetIniPath()
        iniPath = GetIniPath()
        MsgBox "Defaults are stored in: " & iniPath & vbCrLf & _
               "Run Open_AI_Config to edit your settings.", vbInformation, "OllamaLLM"
    End If
End Sub


