VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProviderConfig 
   Caption         =   "Configure AI Providers"
   ClientHeight    =   9120.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7800
   OleObjectBlob   =   "frmProviderConfig.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProviderConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' Current provider being edited
Private mCurrentProviderId As String
Private mIsNewProvider As Boolean

#If VBA7 Then
Private Declare PtrSafe Function GetPrivateProfileStringA Lib "kernel32" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
#Else
Private Declare Function GetPrivateProfileStringA Lib "kernel32" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
#End If

' ============================================================================
' FORM INITIALIZATION
' ============================================================================

Private Sub UserForm_Initialize()
    LoadProviderList
    If cboProvider.ListCount > 0 Then
        cboProvider.ListIndex = 0
    End If
End Sub

' ============================================================================
' PROVIDER LIST MANAGEMENT
' ============================================================================

Private Sub LoadProviderList()
    Dim i As Integer
    Dim count As Long
    
    cboProvider.Clear
    RefreshProviderList
    count = GetProviderListCount()
    
    For i = 0 To count - 1
        cboProvider.AddItem GetProviderNameByIndex(i)
    Next i
End Sub

Private Sub cboProvider_Change()
    If cboProvider.ListIndex >= 0 Then
        mCurrentProviderId = GetProviderIdByIndex(cboProvider.ListIndex)
        mIsNewProvider = False
        LoadProviderDetails mCurrentProviderId
    End If
End Sub

Private Sub LoadProviderDetails(providerId As String)
    txtName.Text = ReadProviderValue(providerId, "name", providerId)
    txtEndpoint.Text = ReadProviderValue(providerId, "endpoint", "")
    txtApiKey.Text = ReadProviderValue(providerId, "api_key", "")
    txtModels.Text = Replace(ReadProviderValue(providerId, "models", ""), ",", vbCrLf)
    txtTemperature.Text = ReadProviderValue(providerId, "temperature", "0.2")
    txtMaxTokens.Text = ReadProviderValue(providerId, "max_tokens", "512")
    txtSystem.Text = ReadProviderValue(providerId, "system", "")
    
    ' Set default model in combo
    LoadDefaultModelCombo providerId
    
    ' Reset status
    lblStatus.Caption = ""
    lblStatus.ForeColor = &H0
End Sub

Private Sub LoadDefaultModelCombo(providerId As String)
    Dim models() As String
    Dim defaultModel As String
    Dim i As Integer
    
    cboDefaultModel.Clear
    models = GetModelList(providerId)
    defaultModel = ReadProviderValue(providerId, "default_model", "")
    
    For i = LBound(models) To UBound(models)
        cboDefaultModel.AddItem Trim$(models(i))
        If LCase$(Trim$(models(i))) = LCase$(defaultModel) Then
            cboDefaultModel.ListIndex = cboDefaultModel.ListCount - 1
        End If
    Next i
    
    ' If no match found, add and select default
    If cboDefaultModel.ListIndex < 0 And Len(defaultModel) > 0 Then
        cboDefaultModel.AddItem defaultModel
        cboDefaultModel.ListIndex = cboDefaultModel.ListCount - 1
    End If
End Sub

Private Function ReadProviderValue(providerId As String, key As String, defaultValue As String) As String
    ReadProviderValue = ReadIniValue("provider." & providerId, key, defaultValue)
End Function

Private Function ReadIniValue(section As String, key As String, defaultValue As String) As String
    Dim buffer As String
    Dim length As Long
    buffer = String$(32767, vbNullChar)
    length = GetPrivateProfileStringA(section, key, defaultValue, buffer, Len(buffer), GetProviderIniPath())
    ReadIniValue = Left$(buffer, length)
End Function

' ============================================================================
' ADD / DELETE PROVIDER
' ============================================================================

Private Sub btnAdd_Click()
    Dim newName As String
    Dim newId As String
    
    newName = InputBox("Enter a name for the new provider:", "Add Provider", "New Provider")
    If Len(Trim$(newName)) = 0 Then Exit Sub
    
    ' Generate ID from name (lowercase, no spaces)
    newId = LCase$(Replace(Trim$(newName), " ", "_"))
    newId = Replace(newId, "(", "")
    newId = Replace(newId, ")", "")
    
    ' Add to provider list
    AddProvider newId
    
    ' Set default values
    SaveProviderConfig newId, newName, "", "", "", "", "0.2", "512", _
        "You are a helpful assistant working inside Microsoft Excel. Return only the final answer with no extra words."
    
    ' Reload list and select new provider
    LoadProviderList
    
    Dim i As Integer
    For i = 0 To cboProvider.ListCount - 1
        If LCase$(GetProviderIdByIndex(i)) = LCase$(newId) Then
            cboProvider.ListIndex = i
            Exit For
        End If
    Next i
    
    mIsNewProvider = True
End Sub

Private Sub btnDelete_Click()
    Dim providerId As String
    Dim response As VbMsgBoxResult
    
    If cboProvider.ListIndex < 0 Then Exit Sub
    
    providerId = GetProviderIdByIndex(cboProvider.ListIndex)
    
    If GetProviderListCount() <= 1 Then
        MsgBox "Cannot delete the last provider. At least one provider must remain.", _
               vbExclamation, "Delete Provider"
        Exit Sub
    End If
    
    response = MsgBox("Are you sure you want to delete '" & txtName.Text & "'?", _
                      vbQuestion + vbYesNo, "Delete Provider")
    
    If response = vbYes Then
        RemoveProvider providerId
        LoadProviderList
        If cboProvider.ListCount > 0 Then
            cboProvider.ListIndex = 0
        End If
    End If
End Sub

' ============================================================================
' TEST CONNECTION
' ============================================================================

Private Sub btnTest_Click()
    Dim result As String
    
    lblStatus.Caption = "Testing connection..."
    lblStatus.ForeColor = &H808000  ' Dark yellow
    DoEvents
    
    result = TestProviderConnection(txtEndpoint.Text, txtApiKey.Text, cboDefaultModel.Text)
    
    If result = "OK" Then
        lblStatus.Caption = "Connected successfully!"
        lblStatus.ForeColor = &H8000&  ' Green
    Else
        lblStatus.Caption = result
        lblStatus.ForeColor = &HFF&  ' Red
    End If
End Sub

' ============================================================================
' REFRESH MODELS
' ============================================================================

Private Sub btnRefreshModels_Click()
    Dim models As String
    
    lblStatus.Caption = "Fetching models..."
    lblStatus.ForeColor = &H808000  ' Dark yellow
    DoEvents
    
    models = FetchModelsFromProvider(txtEndpoint.Text, txtApiKey.Text)
    
    If Len(models) > 0 Then
        txtModels.Text = Replace(models, ",", vbCrLf)
        lblStatus.Caption = "Models loaded successfully!"
        lblStatus.ForeColor = &H8000&  ' Green
        
        ' Refresh default model combo
        Dim modelArr() As String
        modelArr = Split(models, ",")
        cboDefaultModel.Clear
        Dim i As Integer
        For i = LBound(modelArr) To UBound(modelArr)
            cboDefaultModel.AddItem Trim$(modelArr(i))
        Next i
        If cboDefaultModel.ListCount > 0 Then
            cboDefaultModel.ListIndex = 0
        End If
    Else
        lblStatus.Caption = "Failed to fetch models. Check endpoint and API key."
        lblStatus.ForeColor = &HFF&  ' Red
    End If
End Sub

' ============================================================================
' SAVE / CANCEL
' ============================================================================

Private Sub btnSave_Click()
    Dim providerId As String
    Dim models As String
    Dim defaultModel As String
    
    If cboProvider.ListIndex < 0 And Not mIsNewProvider Then
        MsgBox "Please select a provider to save.", vbExclamation, "Save Provider"
        Exit Sub
    End If
    
    providerId = mCurrentProviderId
    
    ' Convert models from multiline to comma-separated
    models = Replace(txtModels.Text, vbCrLf, ",")
    models = Replace(models, vbLf, ",")
    models = Replace(models, vbCr, ",")
    
    ' Clean up multiple commas
    Do While InStr(models, ",,") > 0
        models = Replace(models, ",,", ",")
    Loop
    If Left$(models, 1) = "," Then models = Mid$(models, 2)
    If Right$(models, 1) = "," Then models = Left$(models, Len(models) - 1)
    
    ' Get default model
    If cboDefaultModel.ListIndex >= 0 Then
        defaultModel = cboDefaultModel.Text
    ElseIf cboDefaultModel.ListCount > 0 Then
        defaultModel = cboDefaultModel.List(0)
    Else
        defaultModel = ""
    End If
    
    ' Save configuration
    SaveProviderConfig providerId, _
                       txtName.Text, _
                       txtEndpoint.Text, _
                       txtApiKey.Text, _
                       models, _
                       defaultModel, _
                       txtTemperature.Text, _
                       txtMaxTokens.Text, _
                       txtSystem.Text
    
    ' Update provider name in list if changed
    LoadProviderList
    
    ' Re-select current provider
    Dim i As Integer
    For i = 0 To cboProvider.ListCount - 1
        If LCase$(GetProviderIdByIndex(i)) = LCase$(providerId) Then
            cboProvider.ListIndex = i
            Exit For
        End If
    Next i
    
    lblStatus.Caption = "Provider saved successfully!"
    lblStatus.ForeColor = &H8000&  ' Green
End Sub

Private Sub btnCancel_Click()
    Me.Hide
    Unload Me
End Sub

' ============================================================================
' FORM CLOSE
' ============================================================================

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Just hide, let the calling code handle unload
End Sub
