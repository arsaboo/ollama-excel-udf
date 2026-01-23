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
    Dim http As Object
    Dim status As Long
    Dim body As String
    Dim payload As String
    Dim json As Object
    Dim content As String
    Dim url As String
    Dim system As String

    EnsureIniDefaults

    model = ResolveIniString(model, "model", "qwen3:30b-a3b-instruct-2507-q8_0")
    endpoint = ResolveIniString(endpoint, "endpoint", "http://192.168.2.162:11434/v1/chat/completions")
    api_key = ResolveIniString(api_key, "api_key", "")

    If IsMissing(temperature) Or IsEmpty(temperature) Then
        temperature = ResolveIniDouble("temperature", 0.2)
    End If

    If IsMissing(max_tokens) Or IsEmpty(max_tokens) Then
        max_tokens = ResolveIniLong("max_tokens", 512)
    End If

    system = ResolveIniString("", "system", DefaultSystemPrompt())

    url = NormalizeEndpoint(endpoint)

    On Error GoTo FailSoft

    payload = BuildChatPayload(prompt, model, CDbl(temperature), CLng(max_tokens), system)

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

    If status <> 200 Then
        AI = "Error: HTTP " & status & " - " & http.StatusText & ". Body: " & Left$(body, 500)
        Exit Function
    End If

    Set json = JsonConverter.ParseJson(body)
    On Error Resume Next
    content = json("choices")(1)("message")("content")
    On Error GoTo FailSoft

    If Len(content) = 0 Then
        AI = "Error: Missing content in response. Raw: " & Left$(body, 500)
    Else
        AI = Trim(content)
    End If
    Exit Function

FailSoft:
    AI = "VBA Error #" & Err.Number & ": " & Err.Description
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

    EnsureIniDefaults

    model = ResolveIniString(model, "model", "sonar-pro", "search")
    endpoint = ResolveIniString(endpoint, "endpoint", "https://api.perplexity.ai", "search")
    api_key = ResolveIniString(api_key, "api_key", "", "search")

    isGemini = InStr(1, endpoint, "generativelanguage.googleapis.com", vbTextCompare) > 0
    isResponses = IsResponsesModel(model, endpoint)

    If IsMissing(temperature) Or IsEmpty(temperature) Then
        temperature = ResolveIniDouble("temperature", 0.2, "search")
    End If

    If IsMissing(max_tokens) Or IsEmpty(max_tokens) Then
        max_tokens = ResolveIniLong("max_tokens", 512, "search")
    End If

    system = ResolveIniString("", "system", DefaultSystemPrompt(), "search")

    If isGemini Then
        url = NormalizeGeminiEndpoint(endpoint, model, api_key)
    ElseIf isResponses Then
        url = NormalizeResponsesEndpoint(endpoint)
    Else
        url = NormalizeEndpoint(endpoint)
    End If

    On Error GoTo FailSoft

    If isGemini Then
        payload = BuildGeminiPayload(prompt, system)
    ElseIf isResponses Then
        payload = BuildResponsesPayload(prompt, model, CDbl(temperature), CLng(max_tokens), system)
    Else
        payload = BuildChatPayload(prompt, model, CDbl(temperature), CLng(max_tokens), system)
    End If

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

    If status <> 200 Then
        AI_SEARCH = "Error: HTTP " & status & " - " & http.StatusText & ". Body: " & Left$(body, 500)
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
    On Error GoTo FailSoft

    If Len(content) = 0 Then
        AI_SEARCH = "Error: Missing content in response. Raw: " & Left$(body, 500)
    Else
        AI_SEARCH = Trim(content)
    End If
    Exit Function

FailSoft:
    AI_SEARCH = "VBA Error #" & Err.Number & ": " & Err.Description
End Function

Public Function AI_Version() As String
    AI_Version = "2026-01-23.1"
End Function

' Build OpenAI-compatible payload (uses strongly-typed Dictionary for VBA-JSON)
Private Function BuildChatPayload(prompt As String, _
                                  model As String, _
                                  temperature As Double, _
                                  max_tokens As Long, _
                                  system As String) As String
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


