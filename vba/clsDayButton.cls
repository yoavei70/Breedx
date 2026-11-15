VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDayButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' One day cell of the pop-up calendar in frmSalesEntry.
' Btn.Tag holds the day's date serial; a click reports it back to the form.
'==============================================================================
Option Explicit

Public WithEvents Btn As MSForms.CommandButton
Attribute Btn.VB_VarHelpID = -1
Public Owner As Object      ' frmSalesEntry

Private Sub Btn_Click()
    If Len(Btn.Tag) > 0 Then Owner.CalendarDayPicked CDate(CLng(Btn.Tag))
End Sub
