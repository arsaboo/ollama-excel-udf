Attribute VB_Name = "modAI_Function"
Option Explicit

' === AI worksheet function ===
' NOTE: Requires Tools ? References ? Microsoft Scripting Runtime
Public Function AI(prompt As String, _
                   Optional model As String = "qwen3:30b-a3b-instruct-2507-q8_0", _
                   Optional temperature As Double = 0.2, _
                   Optional max_tokens As Long = 512, _
                   Optional system As String = "", _
                   Optional endpoint As String = "http://192.168.2.162:11434/v1/chat/completions") As String
Attribute AI.VB_Description = "Send a prompt to your Ollama server and return a short, Excel-friendly answer."
Attribute AI.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim http As Object
    Dim status As Long
    Dim body As String
    Dim payload As String
    Dim json As Object
    Dim content As String
    Dim url As String

    If Len(system) = 0 Then
        system = "You are a helpful assistant working inside Microsoft Excel. " & _
                 "Always return only the most concise, direct answer to the user’s question. " & _
                 "Do not include explanations, context, or extra words. " & _
                 "Use plain text only (no Markdown). " & _
                 "If the answer is a single value, output only that value."
    End If

    url = NormalizeEndpoint(endpoint)

    On Error GoTo FailSoft

    payload = BuildChatPayload(prompt, model, temperature, max_tokens, system)

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 30000, 30000, 30000, 120000
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/json"
    http.SetRequestHeader "Accept", "application/json"
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

' Accepts host-only or full path; appends /v1/chat/completions if needed
Private Function NormalizeEndpoint(ByVal e As String) As String
    Dim s As String
    s = Trim(e)
    If Len(s) = 0 Then
        s = "http://127.0.0.1:11434/v1/chat/completions"
    End If
    ' If it ends with /api/chat or /v1/chat/completions, leave as is
    If Right$(s, 14) = "/api/chat" Or Right$(s, 21) = "/v1/chat/completions" Then
        NormalizeEndpoint = s
        Exit Function
    End If
    ' If it looks like just scheme://host[:port] or with trailing slash, append path
    If InStr(1, s, "/v1/chat/completions", vbTextCompare) = 0 And _
       InStr(1, s, "/api/chat", vbTextCompare) = 0 Then
        If Right$(s, 1) = "/" Then
            s = Left$(s, Len(s) - 1)
        End If
        s = s & "/v1/chat/completions"
    End If
    NormalizeEndpoint = s
End Function


