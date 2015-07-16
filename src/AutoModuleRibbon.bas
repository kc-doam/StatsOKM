Attribute VB_Name = "AutoModuleRibbon"
Option Explicit
Option Private Module ' rev.330
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890

Dim Rib As IRibbonUI, ControlTag As String

Private Sub RibbonOnLoad(ByVal ribbon As IRibbonUI) ' customUI.onLoad
  'Shell "Calc.exe", vbNormalFocus
  Set Rib = ribbon
  ' Если необходимо запустить макрос при открытии книги, то использовать:
  Call AutoOpenControls
End Sub

Private Sub GetEnabledMacro(ByVal control As IRibbonControl, ByRef Enabled)
  If ControlTag = "Enable" Then Enabled = True Else _
    Enabled = IIf(control.Tag Like ControlTag, True, False)
End Sub

Private Sub GetVisibleMenu(ByVal control As IRibbonControl, ByRef Visible)
  If ControlTag = "show" Then Visible = True Else _
    Visible = IIf(Not ActiveSheet.ProtectScenarios, True, False)
End Sub

Private Sub LetSheetProtect(ByVal control As IRibbonControl, ByRef CancelDefault)
  CancelDefault = False ' Выполнить команду, затем onAction rev.330
  Call RefreshRibbon(Tag:=ControlTag)
End Sub

Sub DisableAllControls()
  Stop 'Call RefreshRibbon(Tag:="")
End Sub

Sub EnabledAllControls() ' Включить все элементы управления
  Stop 'Call RefreshRibbon(Tag:="*")
End Sub

Sub AutoOpenControls() ' Open #1.5 Отключить элементы управления, кроме группы
  Call RefreshRibbon(Tag:="G" & IIf(ActiveSheet.FilterMode, "*", "0*"))
End Sub

Private Sub RefreshRibbon(ByVal Tag As String) ' Не останавливать!
  ControlTag = Tag
  If Rib Is Nothing Then
    MsgBox "Сохраните/Перезагрузите рабочую книгу " & vbCr & vbCr _
      & "Справка: Возможное решение проблемы смотрите на странице: " & vbCr _
      & "http://www.rondebruin.nl/win/section2.htm", vbCritical, "Ошибка"
  Else
    Rib.Invalidate
  End If
End Sub

Private Sub SetFilter(ByRef control As IRibbonControl) ' rev.330
  With Selection
    If ActiveSheet.AutoFilterMode And Len(.Value) > 0 Then
      If control.ID Like "__Add*" And .Row > 1 Then
        .AutoFilter Field:=.Column, Criteria1:="=" & .Value
      ElseIf control.ID Like "__Clear*" Then
        On Error Resume Next ' На случай, если значение Автофильтра = Пусто
          ActiveSheet.ShowAllData
      End If
    End If: AutoOpenControls
  End With
End Sub