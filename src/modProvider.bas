Attribute VB_Name = "modProvider"
Option Explicit

' ============================================================================
' PROVIDER MANAGEMENT MODULE
' Handles reading/writing provider configuration from INI file
' ============================================================================

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

' Cache for provider list (refreshed on demand)
Private mProviderList() As String
Private mProviderListLoaded As Boolean

' Cache for active providers
Private mActiveProviderId As String
Private mSearchProviderId As String
Private mProviderCacheInitialized As Boolean

' ============================================================================
' INI PATH
' ============================================================================

Public Function GetProviderIniPath() As String
    GetProviderIniPath = Environ$("APPDATA") & "\OllamaLLM\config.ini"
End Function

' ============================================================================
' INITIALIZATION
' ============================================================================

Public Sub InitializeProviderDefaults()
    Dim iniPath As String
    Dim folderPath As String
    Dim fso As Object
    
    iniPath = GetProviderIniPath()
    folderPath = Left$(iniPath, InStrRev(iniPath, "\") - 1)
    
    ' Ensure folder exists
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    ' Initialize with Ollama as default provider if no providers exist
    If Len(ReadIni("providers", "list", "")) = 0 Then
        ' Set provider list
        WriteIni "providers", "list", "ollama"
        
        ' Set active providers
        WriteIni "active", "provider", "ollama"
        WriteIni "active", "search_provider", "ollama"
        
        ' Configure Ollama provider
        WriteIni "provider.ollama", "name", "Ollama (Local)"
        WriteIni "provider.ollama", "endpoint", "http://localhost:11434"
        WriteIni "provider.ollama", "api_key", ""
        WriteIni "provider.ollama", "models", "llama3.1:8b,qwen3:30b,mistral:7b"
        WriteIni "provider.ollama", "default_model", "llama3.1:8b"
        WriteIni "provider.ollama", "temperature", "0.2"
        WriteIni "provider.ollama", "max_tokens", "512"
        WriteIni "provider.ollama", "system", "You are a helpful assistant working inside Microsoft Excel. Return only the final answer with no extra words. Do not include explanations, context, or additional sentences. Use plain text only (no Markdown)."
    End If
    
    ' Clear cache to reload
    mProviderListLoaded = False
    mProviderCacheInitialized = False
End Sub

' ============================================================================
' PROVIDER LIST FUNCTIONS
' ============================================================================

Public Function GetProviderListCount() As Long
    LoadProviderListIfNeeded
    GetProviderListCount = UBound(mProviderList) - LBound(mProviderList) + 1
End Function

Public Function GetProviderIdByIndex(index As Integer) As String
    LoadProviderListIfNeeded
    If index >= LBound(mProviderList) And index <= UBound(mProviderList) Then
        GetProviderIdByIndex = mProviderList(index)
    Else
        GetProviderIdByIndex = ""
    End If
End Function

Public Function GetProviderNameByIndex(index As Integer) As String
    Dim providerId As String
    providerId = GetProviderIdByIndex(index)
    If Len(providerId) > 0 Then
        GetProviderNameByIndex = ReadIni("provider." & providerId, "name", providerId)
    Else
        GetProviderNameByIndex = ""
    End If
End Function

Public Function GetProviderIndexById(providerId As String) As Integer
    Dim i As Integer
    LoadProviderListIfNeeded
    For i = LBound(mProviderList) To UBound(mProviderList)
        If LCase$(mProviderList(i)) = LCase$(providerId) Then
            GetProviderIndexById = i
            Exit Function
        End If
    Next i
    GetProviderIndexById = 0
End Function

Private Sub LoadProviderListIfNeeded()
    Dim listStr As String
    If mProviderListLoaded Then Exit Sub
    
    listStr = ReadIni("providers", "list", "ollama")
    If Len(listStr) > 0 Then
        mProviderList = Split(listStr, ",")
    Else
        ReDim mProviderList(0)
        mProviderList(0) = "ollama"
    End If
    mProviderListLoaded = True
End Sub

Public Sub RefreshProviderList()
    mProviderListLoaded = False
    LoadProviderListIfNeeded
End Sub

' ============================================================================
' ACTIVE PROVIDER FUNCTIONS
' ============================================================================

Public Function GetActiveProviderId() As String
    If Not mProviderCacheInitialized Then
        mActiveProviderId = ReadIni("active", "provider", "ollama")
        mSearchProviderId = ReadIni("active", "search_provider", "ollama")
        mProviderCacheInitialized = True
    End If
    GetActiveProviderId = mActiveProviderId
End Function

Public Function GetActiveProviderIndex() As Integer
    GetActiveProviderIndex = GetProviderIndexById(GetActiveProviderId())
End Function

Public Sub SetActiveProvider(providerId As String)
    mActiveProviderId = providerId
    mProviderCacheInitialized = True
    WriteIni "active", "provider", providerId
End Sub

' ============================================================================
' SEARCH PROVIDER FUNCTIONS
' ============================================================================

Public Function GetSearchProviderId() As String
    If Not mProviderCacheInitialized Then
        mActiveProviderId = ReadIni("active", "provider", "ollama")
        mSearchProviderId = ReadIni("active", "search_provider", "ollama")
        mProviderCacheInitialized = True
    End If
    GetSearchProviderId = mSearchProviderId
End Function

Public Function GetSearchProviderIndex() As Integer
    GetSearchProviderIndex = GetProviderIndexById(GetSearchProviderId())
End Function

Public Sub SetSearchProvider(providerId As String)
    mSearchProviderId = providerId
    mProviderCacheInitialized = True
    WriteIni "active", "search_provider", providerId
End Sub

' ============================================================================
' MODEL LIST FUNCTIONS
' ============================================================================

Public Function GetModelListCount(providerId As String) As Long
    Dim models() As String
    models = GetModelList(providerId)
    GetModelListCount = UBound(models) - LBound(models) + 1
End Function

Public Function GetModelNameByIndex(providerId As String, index As Integer) As String
    Dim models() As String
    models = GetModelList(providerId)
    If index >= LBound(models) And index <= UBound(models) Then
        GetModelNameByIndex = models(index)
    Else
        GetModelNameByIndex = ""
    End If
End Function

Public Function GetModelList(providerId As String) As String()
    Dim modelsStr As String
    Dim raw() As String
    Dim result() As String
    Dim i As Long
    Dim count As Long
    
    modelsStr = ReadIni("provider." & providerId, "models", "")
    If Len(modelsStr) > 0 Then
        raw = Split(modelsStr, ",")
        count = 0
        For i = LBound(raw) To UBound(raw)
            Dim item As String
            item = Trim$(raw(i))
            If Len(item) > 0 Then
                If count = 0 Then
                    ReDim result(0)
                Else
                    ReDim Preserve result(count)
                End If
                result(count) = item
                count = count + 1
            End If
        Next i
        If count = 0 Then
            ReDim result(0)
            result(0) = "(no models)"
        End If
    Else
        ReDim result(0)
        result(0) = "(no models)"
    End If
    GetModelList = result
End Function

Public Function GetActiveModelIndex(providerId As String) As Integer
    Dim defaultModel As String
    Dim models() As String
    Dim i As Integer
    
    defaultModel = ReadIni("provider." & providerId, "default_model", "")
    models = GetModelList(providerId)
    
    For i = LBound(models) To UBound(models)
        If LCase$(models(i)) = LCase$(defaultModel) Then
            GetActiveModelIndex = i
            Exit Function
        End If
    Next i
    GetActiveModelIndex = 0
End Function

Public Sub SetActiveModel(providerId As String, modelName As String)
    WriteIni "provider." & providerId, "default_model", modelName
End Sub

' ============================================================================
' PROVIDER CONFIGURATION FUNCTIONS
' ============================================================================

Public Function GetProviderEndpoint(providerId As String) As String
    GetProviderEndpoint = ReadIni("provider." & providerId, "endpoint", "")
End Function

Public Function GetProviderApiKey(providerId As String) As String
    GetProviderApiKey = ReadIni("provider." & providerId, "api_key", "")
End Function

Public Function GetProviderTemperature(providerId As String) As Double
    Dim tempStr As String
    tempStr = ReadIni("provider." & providerId, "temperature", "0.2")
    On Error Resume Next
    GetProviderTemperature = CDbl(tempStr)
    If Err.Number <> 0 Then GetProviderTemperature = 0.2
    On Error GoTo 0
End Function

Public Function GetProviderMaxTokens(providerId As String) As Long
    Dim tokensStr As String
    tokensStr = ReadIni("provider." & providerId, "max_tokens", "512")
    On Error Resume Next
    GetProviderMaxTokens = CLng(tokensStr)
    If Err.Number <> 0 Then GetProviderMaxTokens = 512
    On Error GoTo 0
End Function

Public Function GetProviderSystem(providerId As String) As String
    GetProviderSystem = ReadIni("provider." & providerId, "system", "You are a helpful assistant.")
End Function

Public Function GetProviderDefaultModel(providerId As String) As String
    GetProviderDefaultModel = ReadIni("provider." & providerId, "default_model", "")
End Function

' ============================================================================
' SAVE PROVIDER CONFIGURATION
' ============================================================================

Public Sub SaveProviderConfig(providerId As String, _
                              providerName As String, _
                              endpoint As String, _
                              apiKey As String, _
                              models As String, _
                              defaultModel As String, _
                              temperature As String, _
                              maxTokens As String, _
                              systemPrompt As String)
    
    WriteIni "provider." & providerId, "name", providerName
    WriteIni "provider." & providerId, "endpoint", endpoint
    WriteIni "provider." & providerId, "api_key", apiKey
    WriteIni "provider." & providerId, "models", models
    WriteIni "provider." & providerId, "default_model", defaultModel
    WriteIni "provider." & providerId, "temperature", temperature
    WriteIni "provider." & providerId, "max_tokens", maxTokens
    WriteIni "provider." & providerId, "system", systemPrompt
End Sub

' ============================================================================
' ADD/REMOVE PROVIDERS
' ============================================================================

Public Sub AddProvider(providerId As String)
    Dim currentList As String
    currentList = ReadIni("providers", "list", "")
    
    ' Check if already exists
    If InStr(1, "," & currentList & ",", "," & providerId & ",", vbTextCompare) > 0 Then
        Exit Sub
    End If
    
    If Len(currentList) > 0 Then
        currentList = currentList & "," & providerId
    Else
        currentList = providerId
    End If
    
    WriteIni "providers", "list", currentList
    mProviderListLoaded = False
End Sub

Public Sub RemoveProvider(providerId As String)
    Dim currentList As String
    Dim providers() As String
    Dim newList As String
    Dim i As Integer
    
    ' Cannot remove if it's the only provider
    If GetProviderListCount() <= 1 Then
        MsgBox "Cannot remove the last provider.", vbExclamation, "AI Tools"
        Exit Sub
    End If
    
    currentList = ReadIni("providers", "list", "")
    providers = Split(currentList, ",")
    
    newList = ""
    For i = LBound(providers) To UBound(providers)
        If LCase$(Trim$(providers(i))) <> LCase$(providerId) Then
            If Len(newList) > 0 Then newList = newList & ","
            newList = newList & Trim$(providers(i))
        End If
    Next i
    
    WriteIni "providers", "list", newList
    
    ' If removed provider was active, switch to first available
    If LCase$(GetActiveProviderId()) = LCase$(providerId) Then
        mProviderListLoaded = False
        SetActiveProvider GetProviderIdByIndex(0)
    End If
    If LCase$(GetSearchProviderId()) = LCase$(providerId) Then
        mProviderListLoaded = False
        SetSearchProvider GetProviderIdByIndex(0)
    End If
    
    mProviderListLoaded = False
End Sub

' ============================================================================
' CONNECTION TEST
' ============================================================================

Public Function TestProviderConnection(endpoint As String, apiKey As String, Optional modelName As String = "") As String
    Dim http As Object
    Dim url As String
    Dim status As Long
    Dim body As String
    Dim isOllama As Boolean
    Dim isOpenAI As Boolean
    Dim isGemini As Boolean
    Dim isPerplexity As Boolean
    Dim isOpenRouter As Boolean
    Dim usePost As Boolean
    Dim payload As String
    
    On Error GoTo TestFailed
    
    ' Determine provider type from endpoint
    isOllama = InStr(1, endpoint, "11434", vbTextCompare) > 0 Or _
               InStr(1, endpoint, "ollama", vbTextCompare) > 0
    isGemini = InStr(1, endpoint, "generativelanguage.googleapis.com", vbTextCompare) > 0
    isPerplexity = InStr(1, endpoint, "api.perplexity.ai", vbTextCompare) > 0
    isOpenRouter = InStr(1, endpoint, "openrouter.ai", vbTextCompare) > 0
    isOpenAI = InStr(1, endpoint, "api.openai.com", vbTextCompare) > 0
    
    ' Build test URL
    url = Trim$(endpoint)
    If Right$(url, 1) = "/" Then url = Left$(url, Len(url) - 1)
    
    If isOllama Then
        url = url & "/api/version"
    ElseIf isGemini Then
        url = url & "/models?key=" & apiKey
    ElseIf isPerplexity Then
        url = url & "/chat/completions"
        usePost = True
        If Len(modelName) = 0 Then modelName = "sonar"
    ElseIf isOpenRouter Then
        If Right$(url, 6) = "/api/v1" Then
            url = url & "/chat/completions"
        ElseIf Right$(url, 3) = "/v1" Then
            url = url & "/chat/completions"
        Else
            url = url & "/api/v1/chat/completions"
        End If
        usePost = True
        If Len(modelName) = 0 Then modelName = "openai/gpt-3.5-turbo"
    ElseIf isOpenAI Then
        url = url & "/v1/models"
    Else
        ' Generic test - try OpenAI-style models endpoint
        url = url & "/v1/models"
    End If
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 5000, 5000, 5000, 10000
    If usePost Then
        http.Open "POST", url, False
        http.SetRequestHeader "Content-Type", "application/json"
    Else
        http.Open "GET", url, False
    End If
    http.SetRequestHeader "Accept", "application/json"
    
    If Len(apiKey) > 0 And Not isGemini Then
        http.SetRequestHeader "Authorization", "Bearer " & apiKey
    End If
    
    If usePost Then
        payload = "{""model"":""" & Replace(modelName, """", "\""") & """,""messages"":[{""role"":""user"",""content"":""ping""}],""max_tokens"":1}"
        http.Send payload
    Else
        http.Send
    End If
    
    status = http.status
    body = http.responseText
    
    If status = 200 Then
        TestProviderConnection = "OK"
    Else
        TestProviderConnection = "Error: HTTP " & status & " - " & http.StatusText & ". Body: " & Left$(body, 500)
    End If
    Exit Function
    
TestFailed:
    TestProviderConnection = "Error: " & Err.Description
End Function

' ============================================================================
' FETCH MODELS FROM PROVIDER
' ============================================================================

Public Function FetchModelsFromProvider(endpoint As String, apiKey As String) As String
    Dim http As Object
    Dim url As String
    Dim status As Long
    Dim body As String
    Dim json As Object
    Dim models As String
    Dim i As Long
    Dim isOllama As Boolean
    Dim isOpenAI As Boolean
    Dim isGemini As Boolean
    
    On Error GoTo FetchFailed
    
    ' Determine provider type from endpoint
    isOllama = InStr(1, endpoint, "11434", vbTextCompare) > 0 Or _
               InStr(1, endpoint, "ollama", vbTextCompare) > 0
    isGemini = InStr(1, endpoint, "generativelanguage.googleapis.com", vbTextCompare) > 0
    isOpenAI = InStr(1, endpoint, "api.openai.com", vbTextCompare) > 0 Or _
               InStr(1, endpoint, "api.perplexity.ai", vbTextCompare) > 0 Or _
               InStr(1, endpoint, "openrouter.ai", vbTextCompare) > 0
    
    ' Build URL
    url = Trim$(endpoint)
    If Right$(url, 1) = "/" Then url = Left$(url, Len(url) - 1)
    
    If isOllama Then
        url = url & "/api/tags"
    ElseIf isGemini Then
        url = url & "/models?key=" & apiKey
    ElseIf isOpenAI Then
        url = url & "/v1/models"
    Else
        url = url & "/v1/models"
    End If
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 5000, 5000, 5000, 15000
    http.Open "GET", url, False
    http.SetRequestHeader "Accept", "application/json"
    
    If Len(apiKey) > 0 And Not isGemini Then
        http.SetRequestHeader "Authorization", "Bearer " & apiKey
    End If
    
    http.Send
    
    status = http.status
    body = http.responseText
    
    If status <> 200 Then
        FetchModelsFromProvider = ""
        Exit Function
    End If
    
    Set json = JsonConverter.ParseJson(body)
    models = ""
    
    On Error Resume Next
    If isOllama Then
        ' Ollama format: {"models": [{"name": "llama3.1:8b", ...}, ...]}
        For i = 1 To json("models").Count
            If Len(models) > 0 Then models = models & ","
            models = models & json("models")(i)("name")
        Next i
    ElseIf isGemini Then
        ' Gemini format: {"models": [{"name": "models/gemini-pro", ...}, ...]}
        For i = 1 To json("models").Count
            If Len(models) > 0 Then models = models & ","
            ' Extract model name from "models/gemini-pro" format
            Dim modelName As String
            modelName = json("models")(i)("name")
            If Left$(modelName, 7) = "models/" Then
                modelName = Mid$(modelName, 8)
            End If
            models = models & modelName
        Next i
    Else
        ' OpenAI format: {"data": [{"id": "gpt-4o", ...}, ...]}
        For i = 1 To json("data").Count
            If Len(models) > 0 Then models = models & ","
            models = models & json("data")(i)("id")
        Next i
    End If
    On Error GoTo FetchFailed
    
    FetchModelsFromProvider = models
    Exit Function
    
FetchFailed:
    FetchModelsFromProvider = ""
End Function

' ============================================================================
' INI HELPERS
' ============================================================================

Private Function ReadIni(section As String, key As String, defaultValue As String) As String
    Dim buffer As String
    Dim length As Long
    buffer = String$(32767, vbNullChar)
    length = GetPrivateProfileString(section, key, defaultValue, buffer, Len(buffer), GetProviderIniPath())
    ReadIni = Left$(buffer, length)
End Function

Private Sub WriteIni(section As String, key As String, value As String)
    WritePrivateProfileString section, key, value, GetProviderIniPath()
End Sub

' ============================================================================
' GET CURRENT PROVIDER SETTINGS (for AI functions)
' ============================================================================

Public Function GetCurrentEndpoint() As String
    GetCurrentEndpoint = GetProviderEndpoint(GetActiveProviderId())
End Function

Public Function GetCurrentApiKey() As String
    GetCurrentApiKey = GetProviderApiKey(GetActiveProviderId())
End Function

Public Function GetCurrentModel() As String
    GetCurrentModel = GetProviderDefaultModel(GetActiveProviderId())
End Function

Public Function GetCurrentTemperature() As Double
    GetCurrentTemperature = GetProviderTemperature(GetActiveProviderId())
End Function

Public Function GetCurrentMaxTokens() As Long
    GetCurrentMaxTokens = GetProviderMaxTokens(GetActiveProviderId())
End Function

Public Function GetCurrentSystem() As String
    GetCurrentSystem = GetProviderSystem(GetActiveProviderId())
End Function

' Search provider settings
Public Function GetSearchEndpoint() As String
    GetSearchEndpoint = GetProviderEndpoint(GetSearchProviderId())
End Function

Public Function GetSearchApiKey() As String
    GetSearchApiKey = GetProviderApiKey(GetSearchProviderId())
End Function

Public Function GetSearchModel() As String
    GetSearchModel = GetProviderDefaultModel(GetSearchProviderId())
End Function

Public Function GetSearchTemperature() As Double
    GetSearchTemperature = GetProviderTemperature(GetSearchProviderId())
End Function

Public Function GetSearchMaxTokens() As Long
    GetSearchMaxTokens = GetProviderMaxTokens(GetSearchProviderId())
End Function

Public Function GetSearchSystem() As String
    GetSearchSystem = GetProviderSystem(GetSearchProviderId())
End Function
