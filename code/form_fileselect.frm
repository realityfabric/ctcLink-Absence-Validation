VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_fileselect 
   Caption         =   "Absence Validation - File Selection"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   9672.001
   OleObjectBlob   =   "form_fileselect.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "form_fileselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Forms")
Option Explicit

Private Sub button_fileselect_abvalidation_Click()
    '@Ignore UseMeaningfulName
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "Select QHC_AB_VALIDATION_ENT_CAL_ERCD"
        If .Show = -1 Then
            Me.textbox_fileselect_abvalidation.Value = .SelectedItems.Item(1)
        End If
    End With
End Sub

Private Sub button_fileselect_jobdata_Click()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "Select QHC_HR_CTC_JOB_DATA"
        If .Show = -1 Then
            Me.textbox_fileselect_jobdata.Value = .SelectedItems.Item(1)
        End If
    End With
End Sub

Private Sub button_run_Click()
    ' Ensure form is filled out
    If Len(Me.textbox_fileselect_abvalidation.Value) = 0 Then
        MsgBox "You Must Select A File for QHC_AB_VALIDATION_ENT_CAL_ERCD." _
        , vbExclamation
    ElseIf Len(Me.textbox_fileselect_jobdata.Value) = 0 Then
        MsgBox "You Must Select A File for QHC_HR_CTC_JOB_DATA." _
        , vbExclamation
    Else
        ' Set session variables
        Main.Sesh.fpathABValidation = Me.textbox_fileselect_abvalidation.Value
        Main.Sesh.fpathJobData = Me.textbox_fileselect_jobdata.Value
        Main.Sesh.FormClosedWithoutRunning = False
        
        Me.Hide
    End If
End Sub


