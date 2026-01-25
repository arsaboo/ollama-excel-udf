Attribute VB_Name = "modRibbon"
Option Explicit

' Ribbon reference for invalidation
Private pRibbon As IRibbonUI

' ============================================================================
' RIBBON INITIALIZATION
' ============================================================================

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set pRibbon = ribbon
    ' Ensure defaults exist
    InitializeProviderDefaults
End Sub

Public Sub RefreshRibbon()
    If Not pRibbon Is Nothing Then
        pRibbon.Invalidate
    End If
End Sub

Public Sub RefreshProviderDropdowns()
    If Not pRibbon Is Nothing Then
        pRibbon.InvalidateControl "ddProvider"
        pRibbon.InvalidateControl "ddModel"
        pRibbon.InvalidateControl "ddSearchProvider"
        pRibbon.InvalidateControl "ddSearchModel"
    End If
End Sub

' ============================================================================
' INSERT FUNCTION CALLBACKS
' ============================================================================

Public Sub OnInsertAI(control As IRibbonControl)
    InsertFormulaAtCursor "=AI("""")"
End Sub

Public Sub OnInsertAISearch(control As IRibbonControl)
    InsertFormulaAtCursor "=AI_SEARCH("""")"
End Sub

Public Sub OnInsertExtract(control As IRibbonControl)
    InsertFormulaAtCursor "=AI_EXTRACT(,"""")"
End Sub

Public Sub OnInsertClassify(control As IRibbonControl)
    InsertFormulaAtCursor "=AI_CLASSIFY(,"""")"
End Sub

Public Sub OnInsertTranslate(control As IRibbonControl)
    InsertFormulaAtCursor "=AI_TRANSLATE(,"""")"
End Sub

Public Sub OnInsertSummarize(control As IRibbonControl)
    InsertFormulaAtCursor "=AI_SUMMARIZE()"
End Sub

Public Sub OnInsertSentiment(control As IRibbonControl)
    InsertFormulaAtCursor "=AI_SENTIMENT()"
End Sub

Public Sub OnInsertFix(control As IRibbonControl)
    InsertFormulaAtCursor "=AI_FIX()"
End Sub

Private Sub InsertFormulaAtCursor(formula As String)
    On Error Resume Next
    ActiveCell.formula = formula
    ' Try to position cursor inside the formula
    Application.SendKeys "{F2}"
    On Error GoTo 0
End Sub

' ============================================================================
' BULK FILL CALLBACK
' ============================================================================

Public Sub OnBulkFill(control As IRibbonControl)
    frmAIBulk.Show
End Sub

' ============================================================================
' PROVIDER DROPDOWN CALLBACKS
' ============================================================================

Public Sub GetProviderCount(control As IRibbonControl, ByRef count)
    count = GetProviderListCount()
End Sub

Public Sub GetProviderLabel(control As IRibbonControl, index As Integer, ByRef label)
    label = GetProviderNameByIndex(index)
End Sub

Public Sub GetProviderSelectedIndex(control As IRibbonControl, ByRef index)
    index = GetActiveProviderIndex()
End Sub

Public Sub OnProviderChange(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    SetActiveProvider GetProviderIdByIndex(selectedIndex)
    ' Refresh model dropdown for new provider
    If Not pRibbon Is Nothing Then
        pRibbon.InvalidateControl "ddModel"
    End If
End Sub

' ============================================================================
' MODEL DROPDOWN CALLBACKS
' ============================================================================

Public Sub GetModelCount(control As IRibbonControl, ByRef count)
    count = GetModelListCount(GetActiveProviderId())
End Sub

Public Sub GetModelLabel(control As IRibbonControl, index As Integer, ByRef label)
    label = GetModelNameByIndex(GetActiveProviderId(), index)
End Sub

Public Sub GetModelSelectedIndex(control As IRibbonControl, ByRef index)
    index = GetActiveModelIndex(GetActiveProviderId())
End Sub

Public Sub OnModelChange(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    SetActiveModel GetActiveProviderId(), GetModelNameByIndex(GetActiveProviderId(), selectedIndex)
End Sub

' ============================================================================
' SEARCH PROVIDER DROPDOWN CALLBACKS
' ============================================================================

Public Sub GetSearchProviderCount(control As IRibbonControl, ByRef count)
    count = GetProviderListCount()
End Sub

Public Sub GetSearchProviderLabel(control As IRibbonControl, index As Integer, ByRef label)
    label = GetProviderNameByIndex(index)
End Sub

Public Sub GetSearchProviderSelectedIndex(control As IRibbonControl, ByRef index)
    index = GetSearchProviderIndex()
End Sub

Public Sub OnSearchProviderChange(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    SetSearchProvider GetProviderIdByIndex(selectedIndex)
    ' Refresh search model dropdown for new provider
    If Not pRibbon Is Nothing Then
        pRibbon.InvalidateControl "ddSearchModel"
    End If
End Sub

' ============================================================================
' SEARCH MODEL DROPDOWN CALLBACKS
' ============================================================================

Public Sub GetSearchModelCount(control As IRibbonControl, ByRef count)
    count = GetModelListCount(GetSearchProviderId())
End Sub

Public Sub GetSearchModelLabel(control As IRibbonControl, index As Integer, ByRef label)
    label = GetModelNameByIndex(GetSearchProviderId(), index)
End Sub

Public Sub GetSearchModelSelectedIndex(control As IRibbonControl, ByRef index)
    index = GetActiveModelIndex(GetSearchProviderId())
End Sub

Public Sub OnSearchModelChange(control As IRibbonControl, selectedId As String, selectedIndex As Integer)
    SetActiveModel GetSearchProviderId(), GetModelNameByIndex(GetSearchProviderId(), selectedIndex)
End Sub

' ============================================================================
' CONFIGURE BUTTONS
' ============================================================================

Public Sub OnConfigureProvider(control As IRibbonControl)
    frmProviderConfig.Show
    RefreshProviderDropdowns
End Sub

Public Sub OnConfigureSearchProvider(control As IRibbonControl)
    frmProviderConfig.Show
    RefreshProviderDropdowns
End Sub

' ============================================================================
' HELP BUTTONS
' ============================================================================

Public Sub OnDocumentation(control As IRibbonControl)
    On Error Resume Next
    ThisWorkbook.FollowHyperlink "https://github.com/your-repo/ollama-excel-udf#readme"
    On Error GoTo 0
End Sub

Public Sub OnOpenConfig(control As IRibbonControl)
    Open_AI_Config
End Sub

Public Sub OnAbout(control As IRibbonControl)
    MsgBox "AI Tools for Excel" & vbCrLf & vbCrLf & _
           "Version: " & AI_Version() & vbCrLf & _
           "Provides AI functions powered by Ollama, OpenAI, Perplexity, and Gemini." & vbCrLf & vbCrLf & _
           "MIT License" & vbCrLf & _
           "JSON parsing via VBA-JSON by Tim Hall", _
           vbInformation, "About AI Tools"
End Sub
