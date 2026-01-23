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
    If Not RegisterVersion(ThisWorkbook.Name & "!AI_Version") Then errs = errs & vbCrLf & " - " & ThisWorkbook.Name & "!AI_Version"
    If Not RegisterVersion("AI_Version") Then errs = errs & vbCrLf & " - AI_Version"
    On Error Resume Next
    Application.Run ThisWorkbook.Name & "!AI_Notify_First_Run"
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
            "system (optional): Default from INI (ai.system) or built-in system prompt.", _
            "endpoint (optional): Default from INI (ai.endpoint) or built-in default.", _
            "api_key (optional): Default from INI (ai.api_key). Sent as Authorization: Bearer <key>." _
        )
    RegisterOne = True
    Exit Function
Fail:
    RegisterOne = False
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



