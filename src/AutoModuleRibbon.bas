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
  If control.Id = "__Costs" Then ' Кнопка «Цены» rev.390
    With ActiveSheet
      If Not (Len(.Cells(ActiveCell.Row, IIf(.CodeName Like "[OQS]?_", _
        5, 10))) > 0 And ActiveCell.Row > 1) Then Enabled = False
    End With
  End If
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
      & "Примечание: Возможное решение проблемы смотрите на странице: " & vbCr _
      & "http://www.rondebruin.nl/win/section2.htm", vbCritical, _
      "Непредвиденная ошибка Ribbon-меню" ' rev.390
  Else
    Rib.Invalidate
  End If
End Sub

Private Sub SetFilter(ByRef control As IRibbonControl) ' rev.330
  With ActiveCell ' Исключаем ошибку при выделении массива rev.370
    If ActiveSheet.AutoFilterMode And Len(.Value) > 0 Then
      If control.Id Like "__Add*" And .Row > 1 Then
        .AutoFilter Field:=.Column, Criteria1:="=" & .Value
      ElseIf control.Id Like "__Clear*" Then
        On Error Resume Next ' На случай, если значение Автофильтра = Пусто
          ActiveSheet.ShowAllData
      End If
    End If: AutoOpenControls
  End With
End Sub

Private Sub ShowCosts(ByRef control As IRibbonControl) ' rev.390
Dim Supplier As String, PartDate As Variant
  With ActiveSheet
    If .CodeName Like "[OQS]?_" Then
      PartDate = .Cells(ActiveCell.Row, 6).FormulaR1C1
      Supplier = .Cells(ActiveCell.Row, 5): CostUpdate Supplier
      .Cells(PartNumRow, 2).NumberFormat = "@" ' Костыль
    Else ' Исходные данные до ФИО = 15
      If .CodeName = Set_arName Then SuppNumRow = ActiveCell.Row
      PartDate = MultidimArr(.Cells(SuppNumRow, 1).Resize(, 15).Value, 1)
      Supplier = PartDate(10)
    End If: PartDate = GetDateAndCosts(.Name, PartDate)
  End With
  If IsArray(PartDate) Then
    MsgBox "Цены '" & Supplier & "' с " & CDate(PartDate(1)) & vbCr _
      & "Группы 0-2: " & PartDate(2) & " руб; " & PartDate(3) & " руб; " & PartDate(4) & " руб " _
      & vbCr & "Актуализация: " & PartDate(5) & " руб " & vbCr _
      & "НУМ-ы 0-2: " & PartDate(6) & " руб; " & PartDate(7) & " руб; " & PartDate(8) & " руб " _
      & vbCr & "НАШ-и 1-2: " & PartDate(9) & " руб; " & PartDate(10) & " руб "
  End If
End Sub
