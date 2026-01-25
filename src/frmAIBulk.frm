VERSION 5.00
Begin VB.UserForm frmAIBulk
   Caption         =   "AI Bulk Fill"
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblPrompt
      Caption         =   "Global prompt"
      Height          =   255
      Left            =   180
      Top             =   180
      Width           =   7200
   End
   Begin MSForms.TextBox txtPrompt
      Height          =   2100
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      Top             =   480
      Width           =   7200
   End
   Begin MSForms.Label lblMode
      Caption         =   "Mode"
      Height          =   255
      Left            =   180
      Top             =   2820
      Width           =   600
   End
   Begin MSForms.OptionButton optLocal
      Caption         =   "Local"
      Height          =   255
      Left            =   900
      Top             =   2800
      Value           =   -1  'True
      Width           =   900
   End
   Begin MSForms.OptionButton optSearch
      Caption         =   "Search"
      Height          =   255
      Left            =   1920
      Top             =   2800
      Width           =   1200
   End
   Begin MSForms.CommandButton btnRun
      Caption         =   "Run"
      Height          =   360
      Left            =   180
      Top             =   3360
      Width           =   1200
   End
   Begin MSForms.CommandButton btnCancel
      Caption         =   "Cancel"
      Height          =   360
      Left            =   1500
      Top             =   3360
      Width           =   1200
   End
   Begin MSForms.CommandButton btnClose
      Caption         =   "Close"
      Height          =   360
      Left            =   2820
      Top             =   3360
      Width           =   1200
   End
   Begin MSForms.Label lblStatus
      Caption         =   "Ready. Select a cell in your table before running."
      Height          =   600
      Left            =   180
      Top             =   3900
      Width           =   7200
   End
End
Attribute VB_Name = "frmAIBulk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Cancelled As Boolean

Private Sub UserForm_Initialize()
    Cancelled = False
End Sub

Private Sub btnRun_Click()
    Cancelled = False
    UpdateStatus "Running..."
    DoEvents
    modAI_Bulk.RunBulkFill Me
End Sub

Private Sub btnCancel_Click()
    Cancelled = True
    UpdateStatus "Cancelling..."
End Sub

Private Sub btnClose_Click()
    Me.Hide
    Unload Me
End Sub

Public Function PromptText() As String
    PromptText = Trim$(txtPrompt.Text)
End Function

Public Function IsSearchMode() As Boolean
    IsSearchMode = CBool(optSearch.Value)
End Function

Public Sub UpdateStatus(ByVal message As String)
    lblStatus.Caption = message
End Sub
