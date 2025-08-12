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

    On Error Resume Next
    If wasAddin Then
        Windows(ThisWorkbook.Name).Visible = False
        ThisWorkbook.IsAddin = True
    End If
    On Error GoTo 0

    If Len(errs) > 0 Then
        MsgBox "Failed to register tooltips for:" & errs, vbExclamation, "AI() tooltip registration"
    End If
End Sub

Private Function RegisterOne(ByVal macroName As String) As Boolean
    On Error GoTo Fail
    Application.MacroOptions _
        Macro:=macroName, _
        Description:="Send a prompt to your Ollama server and return a short, Excel-friendly answer.", _
        Category:="AI Helpers", _
        ArgumentDescriptions:=Array( _
            "prompt (required): Your question or instruction. Plain text.", _
            "model (optional, default=qwen3:30b-a3b-instruct-2507-q8_0): Exact model on the server (see ollama list).", _
            "temperature (optional, default=0.2): 0.0–1.0. Lower = more deterministic.", _
            "max_tokens (optional, default=512): Maximum response length.", _
            "system (optional): System prompt. Leave blank for concise, single-value answers.", _
            "endpoint (optional, default=http://192.168.2.162:11434/v1/chat/completions): " & _
                "Full API URL or just host:port; host-only will auto-append /v1/chat/completions." _
        )
    RegisterOne = True
    Exit Function
Fail:
    RegisterOne = False
End Function


