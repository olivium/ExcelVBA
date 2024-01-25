VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReport 
   Caption         =   "Welcome to the Report Form"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAddWorksheet_Click()
    ' variable for the tryAgain logic below
    Dim tryAgain As Integer
 
    ' if the user leaves the box empty go to the error handler section
    On Error GoTo errHandler
    
    Worksheets.Add before:=Worksheets(1)
        
    ActiveSheet.Name = InputBox("Please enter a new worksheet name")
    
    ' if no error skip the errHandler
    Exit Sub
    
    ' error handler section
errHandler:
    ' prompt the user to try again
    tryAgain = MsgBox("Invalid Worksheet Name", vbYesNo)
    
    ' if the user picked yes
    If tryAgain = 6 Then
        're-run the procedure
        btnAddWorksheet_Click
    Else
        ' turn off the alert of deleting the worksheet if the user selects no
        Application.DisplayAlerts = False
        
        ' delete the invalid worksheet
        ActiveSheet.Delete
    End If
    
    ActiveSheet.Name = InputBox("Please enter a new worksheet name")
End Sub

Private Sub btnRunReport_Click()
    LoopYearlyReport
End Sub

Private Sub cboWhichSheet_Change()
    Worksheets(Me.cboWhichSheet.Value).Select
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    i = 1
    
    Do While i <= Worksheets.Count
        Me.cboWhichSheet.AddItem Worksheets(i).Name
        i = i + 1
    Loop
End Sub
