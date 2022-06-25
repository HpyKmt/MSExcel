VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHideColumns 
   Caption         =   "Hide Columns"
   ClientHeight    =   1350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6240
   OleObjectBlob   =   "frmHideColumns.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHideColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit





Private Sub ButtonHideColumns_Click()
    Dim LastColumn As Long, TargetRow As Long, CurColumn As Long
    Const FIRST_COLUMN = 1
    
    ' Get the last column in the active sheet.
    LastColumn = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
    
    ' Get the target row where column values are evaluated
    TargetRow = CLng(TextBoxRowNo.Value)
    
    ' Loop through columns from the end
    For CurColumn = LastColumn To FIRST_COLUMN Step -1
    
        ' If the specified expression is detected, hide the entire column
        If InStr(1, Cells(TargetRow, CurColumn), TextToHide.Value) > 0 Then
            Columns(CurColumn).EntireColumn.Hidden = True
        End If
    
    Next CurColumn
    
End Sub


Private Sub TextBoxRowNo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ' Ensure that input value is a positive integer
    If Not IsNumeric(TextBoxRowNo.Value) Then
        MsgBox "ERROR: Input a positive integer!", vbCritical
        Cancel = True
    End If
End Sub
