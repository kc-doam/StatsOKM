Attribute VB_Name = "AutoModuleRibbon"
Option Explicit
'12345678901234567890123456789012345678901234567890123456789012345678901234567890

Dim Rib As IRibbonUI
Public ControlTag As String

Sub RibbonOnLoad(ByVal ribbon As IRibbonUI) ' Callback for customUI.onLoad
  'Shell "Calc.exe", vbNormalFocus
  Set Rib = ribbon
  ' Если необходимо запустить макрос при открытии книги, то использовать вроде:
  Call EnabledAllControls 'AutoOpenControls
End Sub

Sub GetEnabledMacro(ByVal control As IRibbonControl, ByRef Enabled)
  If ControlTag = "Enable" Then ' If control.Tag Like "*Enable" Then
    Enabled = True
  Else
    If control.Tag Like ControlTag Then Enabled = True Else Enabled = False
  End If
End Sub

Sub RefreshRibbon(ByVal Tag As String) ' Не останавливать!
  ControlTag = Tag
  If Rib Is Nothing Then
    MsgBox "Сохраните/Перезагрузите рабочую книгу " & vbCr & vbCr _
      & "Справка: Возможное решение проблемы смотрите на странице: " & vbCr _
      & "http://www.rondebruin.nl/win/s2/win013.htm", vbCritical, "Ошибка"
  Else
    Rib.Invalidate
  End If
End Sub

Sub DisableAllControls()
  Call RefreshRibbon(Tag:="")
End Sub

Sub EnabledAllControls() ' Включить все элементы управления
  Call RefreshRibbon(Tag:="*")
End Sub

Sub AutoOpenControls() ' Отключить все элементы управления, кроме
  Call RefreshRibbon(Tag:="G1*")
End Sub

Private Sub AddFilter(ByRef control As IRibbonControl)
  If ActiveSheet.AutoFilterMode And Len(Selection.Value) > 0 _
  And Selection.Row > 1 Then _
    Selection.AutoFilter Field:=Selection.Column, Criteria1:=Selection.Value
End Sub