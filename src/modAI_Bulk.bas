Attribute VB_Name = "modAI_Bulk"
Option Explicit

Public Const AI_BULK_HOTKEY As String = "^+A"

Public Sub Show_AI_Form()
    frmAIBulk.Show
End Sub

Public Sub Register_AI_Hotkey()
    On Error Resume Next
    Application.OnKey AI_BULK_HOTKEY, "Show_AI_Form"
    On Error GoTo 0
End Sub

Public Sub RunBulkFill(ByVal ui As frmAIBulk)
    Dim region As Range
    Dim promptRow As Range
    Dim headerRow As Range
    Dim dataRange As Range
    Dim rowCount As Long
    Dim colCount As Long
    Dim inputCols As Collection
    Dim outputCols As Collection
    Dim r As Long
    Dim c As Long
    Dim colIndex As Variant
    Dim prompt As String
    Dim columnPrompt As String
    Dim headerText As String
    Dim contextText As String
    Dim cellValue As String
    Dim totalCells As Long
    Dim doneCells As Long
    Dim isSearch As Boolean
    Dim hasPromptRow As Boolean
    Dim minRows As Long
    Dim prevCalc As XlCalculation

    If Len(ui.PromptText()) = 0 Then
        MsgBox "Global prompt is required.", vbExclamation, "AI Bulk Fill"
        Exit Sub
    End If

    Set region = GetCurrentRegionSafe()
    If region Is Nothing Then
        MsgBox "Select a cell in a table before running.", vbExclamation, "AI Bulk Fill"
        Exit Sub
    End If

    ' Determine mode and minimum rows required
    hasPromptRow = ui.HasPromptRow()
    If hasPromptRow Then
        minRows = 3  ' prompt + header + at least 1 data row
    Else
        minRows = 2  ' header + at least 1 data row
    End If

    If region.Rows.Count < minRows Then
        If hasPromptRow Then
            MsgBox "The selected table needs a prompt row, header row, and at least one data row.", _
                   vbExclamation, "AI Bulk Fill"
        Else
            MsgBox "The selected table needs a header row and at least one data row.", _
                   vbExclamation, "AI Bulk Fill"
        End If
        Exit Sub
    End If

    rowCount = region.Rows.Count
    colCount = region.Columns.Count

    ' Set row references based on mode
    If hasPromptRow Then
        Set promptRow = region.Rows(1)
        Set headerRow = region.Rows(2)
        Set dataRange = region.Offset(2, 0).Resize(rowCount - 2, colCount)
    Else
        Set promptRow = Nothing
        Set headerRow = region.Rows(1)
        Set dataRange = region.Offset(1, 0).Resize(rowCount - 1, colCount)
    End If

    Set inputCols = New Collection
    Set outputCols = New Collection

    For c = 1 To colCount
        headerText = Trim$(CStr(headerRow.Cells(1, c).Value))
        If HasDataInColumn(dataRange, c) Then
            inputCols.Add c
        ElseIf Len(headerText) > 0 Then
            outputCols.Add c
        End If
    Next c

    If outputCols.Count = 0 Then
        MsgBox "No output columns detected. Add headers to empty columns and try again.", vbExclamation, "AI Bulk Fill"
        Exit Sub
    End If

    isSearch = ui.IsSearchMode()
    totalCells = dataRange.Rows.Count * outputCols.Count
    doneCells = 0

    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    For r = 1 To dataRange.Rows.Count
        If ui.Cancelled Then Exit For
        contextText = BuildRowContext(headerRow, dataRange.Rows(r), inputCols)
        For Each colIndex In outputCols
            If ui.Cancelled Then Exit For
            c = CLng(colIndex)
            headerText = Trim$(CStr(headerRow.Cells(1, c).Value))
            
            ' Get column-specific prompt if available
            columnPrompt = ""
            If Not promptRow Is Nothing Then
                columnPrompt = Trim$(CStr(promptRow.Cells(1, c).Value))
            End If
            
            prompt = BuildPrompt(columnPrompt, ui.PromptText(), headerText, contextText)
            ui.UpdateStatus "Running... " & (doneCells + 1) & " of " & totalCells
            If isSearch Then
                cellValue = AI_SEARCH(prompt)
            Else
                cellValue = AI(prompt)
            End If
            dataRange.Cells(r, c).Value = cellValue
            doneCells = doneCells + 1
        Next colIndex
    Next r

CleanExit:
    Application.Calculation = prevCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    ui.UpdateStatus "Error: " & Err.Description
    Resume CleanExit
End Sub

Private Function GetCurrentRegionSafe() As Range
    Dim region As Range
    On Error Resume Next
    Set region = ActiveCell.CurrentRegion
    On Error GoTo 0

    If region Is Nothing Then
        Set GetCurrentRegionSafe = Nothing
        Exit Function
    End If

    If region.Cells.Count = 1 Then
        If Len(Trim$(CStr(region.Cells(1, 1).Value))) = 0 Then
            Set GetCurrentRegionSafe = Nothing
        Else
            Set GetCurrentRegionSafe = region
        End If
    Else
        Set GetCurrentRegionSafe = region
    End If
End Function

Private Function HasDataInColumn(ByVal dataRange As Range, ByVal colIndex As Long) As Boolean
    Dim r As Long
    Dim valueText As String

    For r = 1 To dataRange.Rows.Count
        valueText = Trim$(CStr(dataRange.Cells(r, colIndex).Value))
        If Len(valueText) > 0 Then
            HasDataInColumn = True
            Exit Function
        End If
    Next r

    HasDataInColumn = False
End Function

Private Function BuildRowContext(ByVal headerRow As Range, ByVal dataRow As Range, ByVal inputCols As Collection) As String
    Dim parts As Collection
    Dim idx As Variant
    Dim headerText As String
    Dim valueText As String
    Dim result As String

    Set parts = New Collection
    For Each idx In inputCols
        headerText = Trim$(CStr(headerRow.Cells(1, idx).Value))
        valueText = Trim$(CStr(dataRow.Cells(1, idx).Value))
        If Len(valueText) > 0 Then
            If Len(headerText) > 0 Then
                parts.Add headerText & ": " & valueText
            Else
                parts.Add valueText
            End If
        End If
    Next idx

    result = JoinCollection(parts, vbCrLf)
    BuildRowContext = result
End Function

Private Function BuildPrompt(ByVal columnPrompt As String, _
                             ByVal globalPrompt As String, _
                             ByVal headerText As String, _
                             ByVal contextText As String) As String
    Dim prompt As String
    Dim basePrompt As String
    Dim outputRule As String

    outputRule = "Return ONLY the value. No labels, no prefixes, no extra text. " & _
                 "Just the raw answer (text, word, or number as appropriate)."

    ' Use column-specific prompt if available, else global
    If Len(columnPrompt) > 0 Then
        basePrompt = columnPrompt
    Else
        basePrompt = globalPrompt
    End If

    prompt = Trim$(basePrompt)
    If Len(headerText) > 0 Then
        prompt = prompt & vbCrLf & "Target column: " & headerText
    End If
    If Len(contextText) > 0 Then
        prompt = prompt & vbCrLf & "Row data:" & vbCrLf & contextText
    End If
    prompt = prompt & vbCrLf & vbCrLf & outputRule

    BuildPrompt = Trim$(prompt)
End Function

Private Function JoinCollection(ByVal items As Collection, ByVal separator As String) As String
    Dim i As Long
    Dim output As String

    For i = 1 To items.Count
        If i = 1 Then
            output = CStr(items(i))
        Else
            output = output & separator & CStr(items(i))
        End If
    Next i

    JoinCollection = output
End Function
