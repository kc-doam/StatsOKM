VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890

Private XLApp As cExcelEvents

Private Sub Workbook_Open() ' ВАЖНО! Open #1.1
  If ThisWb Is Nothing Then Set ThisWb = ThisWorkbook ' Приложение rev.410
  If XLApp Is Nothing Then Set XLApp = New cExcelEvents
End Sub
