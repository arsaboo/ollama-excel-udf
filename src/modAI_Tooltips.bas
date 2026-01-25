Attribute VB_Name = "modAI_Tooltips"
Option Explicit

' Call this from ThisWorkbook.Workbook_Open
Public Sub Install_AI_Tooltips()
    Dim wasAddin As Boolean
    Dim errs As String

    wasAddin = ThisWorkbook.IsAddin
    ThisWorkbook.IsAddin = False
    On Error Resume Next
    Windows(ThisWorkbook.Name).Visible = True
    On Error GoTo 0

    errs = ""
    If Not RegisterOne(ThisWorkbook.Name & "!AI") Then errs = errs & vbCrLf & " - " & ThisWorkbook.Name & "!AI"
    If Not RegisterOne("AI") Then errs = errs & vbCrLf & " - AI"
    If Not RegisterSearch(ThisWorkbook.Name & "!AI_SEARCH") Then errs = errs & vbCrLf & " - " & ThisWorkbook.Name & "!AI_SEARCH"
    If Not RegisterSearch("AI_SEARCH") Then errs = errs & vbCrLf & " - AI_SEARCH"
    If Not RegisterVersion(ThisWorkbook.Name & "!AI_Version") Then errs = errs & vbCrLf & " - " & ThisWorkbook.Name & "!AI_Version"
    If Not RegisterVersion("AI_Version") Then errs = errs & vbCrLf & " - AI_Version"
    
    ' Register new UDF functions
    If Not RegisterExtract(ThisWorkbook.Name & "!AI_EXTRACT") Then errs = errs & vbCrLf & " - " & ThisWorkbook.Name & "!AI_EXTRACT"
    If Not RegisterExtract("AI_EXTRACT") Then errs = errs & vbCrLf & " - AI_EXTRACT"
    If Not RegisterClassify(ThisWorkbook.Name & "!AI_CLASSIFY") Then errs = errs & vbCrLf & " - " & ThisWorkbook.Name & "!AI_CLASSIFY"
    If Not RegisterClassify("AI_CLASSIFY") Then errs = errs & vbCrLf & " - AI_CLASSIFY"
    If Not RegisterTranslate(ThisWorkbook.Name & "!AI_TRANSLATE") Then errs = errs & vbCrLf & " - " & ThisWorkbook.Name & "!AI_TRANSLATE"
    If Not RegisterTranslate("AI_TRANSLATE") Then errs = errs & vbCrLf & " - AI_TRANSLATE"
    If Not RegisterSummarize(ThisWorkbook.Name & "!AI_SUMMARIZE") Then errs = errs & vbCrLf & " - " & ThisWorkbook.Name & "!AI_SUMMARIZE"
    If Not RegisterSummarize("AI_SUMMARIZE") Then errs = errs & vbCrLf & " - AI_SUMMARIZE"
    If Not RegisterSentiment(ThisWorkbook.Name & "!AI_SENTIMENT") Then errs = errs & vbCrLf & " - " & ThisWorkbook.Name & "!AI_SENTIMENT"
    If Not RegisterSentiment("AI_SENTIMENT") Then errs = errs & vbCrLf & " - AI_SENTIMENT"
    If Not RegisterFix(ThisWorkbook.Name & "!AI_FIX") Then errs = errs & vbCrLf & " - " & ThisWorkbook.Name & "!AI_FIX"
    If Not RegisterFix("AI_FIX") Then errs = errs & vbCrLf & " - AI_FIX"
    
    On Error Resume Next
    Application.Run ThisWorkbook.Name & "!AI_Notify_First_Run"
    Application.Run ThisWorkbook.Name & "!Register_AI_Hotkey"
    On Error GoTo 0

    On Error Resume Next
    If wasAddin Then
        Windows(ThisWorkbook.Name).Visible = False
        ThisWorkbook.IsAddin = True
    End If
    On Error GoTo 0

    If Len(errs) > 0 Then
        MsgBox "Failed to register tooltips for:" & errs, vbExclamation, "AI() tooltip registration"
    End If

    ThisWorkbook.Saved = True
End Sub

Private Function RegisterOne(ByVal macroName As String) As Boolean
    On Error GoTo Fail
    Application.MacroOptions _
        Macro:=macroName, _
        Description:="Send a prompt to your Ollama server and return a short, Excel-friendly answer.", _
        Category:="AI Helpers", _
        ArgumentDescriptions:=Array( _
            "prompt (required): Your question or instruction. Plain text.", _
            "model (optional): Default from INI (ai.model) or built-in default.", _
            "temperature (optional): Default from INI (ai.temperature) or 0.2.", _
            "max_tokens (optional): Default from INI (ai.max_tokens) or 512.", _
            "endpoint (optional): Default from INI (ai.endpoint) or built-in default.", _
            "api_key (optional): Default from INI (ai.api_key). Sent as Authorization: Bearer <key>." _
        )
    RegisterOne = True
    Exit Function
Fail:
    RegisterOne = False
End Function

Private Function RegisterSearch(ByVal macroName As String) As Boolean
    On Error GoTo Fail
    Application.MacroOptions _
        Macro:=macroName, _
        Description:="Send a prompt to your search-enabled AI provider and return a short, Excel-friendly answer.", _
        Category:="AI Helpers", _
        ArgumentDescriptions:=Array( _
            "prompt (required): Your question or instruction. Plain text.", _
            "model (optional): Default from INI (search.model) or built-in default.", _
            "temperature (optional): Default from INI (search.temperature) or 0.2.", _
            "max_tokens (optional): Default from INI (search.max_tokens) or 512.", _
            "endpoint (optional): Default from INI (search.endpoint) or built-in default.", _
            "api_key (optional): Default from INI (search.api_key). Sent as Authorization: Bearer <key> for OpenAI-compatible providers." _
        )
    RegisterSearch = True
    Exit Function
Fail:
    RegisterSearch = False
End Function

Private Function RegisterVersion(ByVal macroName As String) As Boolean
    On Error GoTo Fail
    Application.MacroOptions _
        Macro:=macroName, _
        Description:="Return the installed add-in version string.", _
        Category:="AI Helpers"
    RegisterVersion = True
    Exit Function
Fail:
    RegisterVersion = False
End Function

Private Function RegisterExtract(ByVal macroName As String) As Boolean
    On Error GoTo Fail
    Application.MacroOptions _
        Macro:=macroName, _
        Description:="Extract a specific field from text using AI.", _
        Category:="AI Helpers", _
        ArgumentDescriptions:=Array( _
            "text (required): The text to extract from.", _
            "field (required): The field to extract (e.g., ""email"", ""phone"", ""date"").", _
            "model (optional): Default from INI (ai.model) or built-in default.", _
            "temperature (optional): Default from INI (ai.temperature) or 0.2.", _
            "max_tokens (optional): Default from INI (ai.max_tokens) or 512.", _
            "endpoint (optional): Default from INI (ai.endpoint) or built-in default.", _
            "api_key (optional): Default from INI (ai.api_key)." _
        )
    RegisterExtract = True
    Exit Function
Fail:
    RegisterExtract = False
End Function

Private Function RegisterClassify(ByVal macroName As String) As Boolean
    On Error GoTo Fail
    Application.MacroOptions _
        Macro:=macroName, _
        Description:="Classify text into one of the provided categories.", _
        Category:="AI Helpers", _
        ArgumentDescriptions:=Array( _
            "text (required): The text to classify.", _
            "categories (required): Comma-separated string or cell range of categories.", _
            "model (optional): Default from INI (ai.model) or built-in default.", _
            "temperature (optional): Default from INI (ai.temperature) or 0.2.", _
            "max_tokens (optional): Default from INI (ai.max_tokens) or 512.", _
            "endpoint (optional): Default from INI (ai.endpoint) or built-in default.", _
            "api_key (optional): Default from INI (ai.api_key)." _
        )
    RegisterClassify = True
    Exit Function
Fail:
    RegisterClassify = False
End Function

Private Function RegisterTranslate(ByVal macroName As String) As Boolean
    On Error GoTo Fail
    Application.MacroOptions _
        Macro:=macroName, _
        Description:="Translate text to target language.", _
        Category:="AI Helpers", _
        ArgumentDescriptions:=Array( _
            "text (required): The text to translate.", _
            "targetLang (required): Target language (e.g., ""Spanish"", ""French"", ""German"").", _
            "model (optional): Default from INI (ai.model) or built-in default.", _
            "temperature (optional): Default from INI (ai.temperature) or 0.2.", _
            "max_tokens (optional): Default from INI (ai.max_tokens) or 512.", _
            "endpoint (optional): Default from INI (ai.endpoint) or built-in default.", _
            "api_key (optional): Default from INI (ai.api_key)." _
        )
    RegisterTranslate = True
    Exit Function
Fail:
    RegisterTranslate = False
End Function

Private Function RegisterSummarize(ByVal macroName As String) As Boolean
    On Error GoTo Fail
    Application.MacroOptions _
        Macro:=macroName, _
        Description:="Summarize text to specified word count.", _
        Category:="AI Helpers", _
        ArgumentDescriptions:=Array( _
            "text (required): The text to summarize.", _
            "maxWords (optional): Maximum words in summary. Default: 50.", _
            "model (optional): Default from INI (ai.model) or built-in default.", _
            "temperature (optional): Default from INI (ai.temperature) or 0.2.", _
            "max_tokens (optional): Default from INI (ai.max_tokens) or 512.", _
            "endpoint (optional): Default from INI (ai.endpoint) or built-in default.", _
            "api_key (optional): Default from INI (ai.api_key)." _
        )
    RegisterSummarize = True
    Exit Function
Fail:
    RegisterSummarize = False
End Function

Private Function RegisterSentiment(ByVal macroName As String) As Boolean
    On Error GoTo Fail
    Application.MacroOptions _
        Macro:=macroName, _
        Description:="Analyze sentiment of text. Returns: Positive, Negative, or Neutral.", _
        Category:="AI Helpers", _
        ArgumentDescriptions:=Array( _
            "text (required): The text to analyze.", _
            "model (optional): Default from INI (ai.model) or built-in default.", _
            "temperature (optional): Default from INI (ai.temperature) or 0.2.", _
            "max_tokens (optional): Default from INI (ai.max_tokens) or 512.", _
            "endpoint (optional): Default from INI (ai.endpoint) or built-in default.", _
            "api_key (optional): Default from INI (ai.api_key)." _
        )
    RegisterSentiment = True
    Exit Function
Fail:
    RegisterSentiment = False
End Function

Private Function RegisterFix(ByVal macroName As String) As Boolean
    On Error GoTo Fail
    Application.MacroOptions _
        Macro:=macroName, _
        Description:="Fix grammar, spelling, and formatting issues in text.", _
        Category:="AI Helpers", _
        ArgumentDescriptions:=Array( _
            "text (required): The text to fix.", _
            "rules (optional): Additional rules (e.g., ""use formal tone"", ""British English"").", _
            "model (optional): Default from INI (ai.model) or built-in default.", _
            "temperature (optional): Default from INI (ai.temperature) or 0.2.", _
            "max_tokens (optional): Default from INI (ai.max_tokens) or 512.", _
            "endpoint (optional): Default from INI (ai.endpoint) or built-in default.", _
            "api_key (optional): Default from INI (ai.api_key)." _
        )
    RegisterFix = True
    Exit Function
Fail:
    RegisterFix = False
End Function


