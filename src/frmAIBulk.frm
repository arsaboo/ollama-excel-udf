VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAIBulk 
   Caption         =   "AI Agent"
   ClientHeight    =   5210
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7190
   OleObjectBlob   =   "frmAIBulk.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAIBulk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Public Cancelled As Boolean
Private mIsRunning As Boolean

Private Sub UserForm_Initialize()
    Cancelled = False
    mIsRunning = False
    chkSearch.Value = False
    chkPromptRow.Value = False
    btnClose.Caption = "Close"
    UpdateStatus "Ready. Select a cell in your table before running."
End Sub

Private Sub btnRun_Click()
    If mIsRunning Then Exit Sub
    Cancelled = False
    StartRun
    modAI_Bulk.RunBulkFill Me
    EndRun
End Sub

Private Sub btnClose_Click()
    If mIsRunning Then
        Cancelled = True
        UpdateStatus "Stopping..."
    Else
        Me.Hide
        Unload Me
    End If
End Sub

Public Sub StartRun()
    mIsRunning = True
    Cancelled = False
    btnClose.Caption = "Stop"
    btnRun.Enabled = False
    UpdateStatus "Running..."
    DoEvents
End Sub

Public Sub EndRun()
    mIsRunning = False
    btnClose.Caption = "Close"
    btnRun.Enabled = True
    If Cancelled Then
        UpdateStatus "Cancelled."
    Else
        UpdateStatus "Done."
    End If
    DoEvents
End Sub

Public Function IsRunning() As Boolean
    IsRunning = mIsRunning
End Function

Public Function PromptText() As String
    PromptText = Trim$(txtPrompt.Text)
End Function

Public Function IsSearchMode() As Boolean
    IsSearchMode = CBool(chkSearch.Value)
End Function

Public Function HasPromptRow() As Boolean
    HasPromptRow = CBool(chkPromptRow.Value)
End Function

Public Sub UpdateStatus(ByVal message As String)
    lblStatus.Caption = message
    DoEvents
End Sub
