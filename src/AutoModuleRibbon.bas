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
      If Not (Len(.Cells(ActiveCell.Row, IIf(.CodeName Like SHEETS_ALL, _
        5, 10))) > 0 And ActiveCell.Row > 1) Or .CodeName = "OE_" _
        Or .CodeName = "QT_" Then Enabled = False ' rev.420
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

Sub EnabledAllControls() ' Включить все элементы управления rev.410
  'Call RefreshRibbon(Tag:="*")
  Call RefreshRibbon(Tag:="G" & IIf(ActiveSheet.FilterMode, "*", "0*"))
End Sub

Sub AutoOpenControls() ' Open #1.5 Отключить элементы управления, кроме группы
  ' ВАЖНО! Значения "TimeSerial(0, 0, 1) / &H" подбираются опытным путём
  If Rib Is Nothing Then Exit Sub ' rev.420
  If cnfRenew = "Ф/Л" Or Len(cnfRenew) = 0 Then _
    Application.OnTime Now + TimeSerial(0, 0, 1) / &H8, "EnabledAllControls" _
  Else EnabledAllControls
End Sub

Private Sub RefreshRibbon(ByVal Tag As String) ' Не останавливать! rev.410
  ControlTag = Tag
  If Rib Is Nothing Then
    UnprotectSheet ThisWb.ActiveSheet ' rev.400
    ThisWb.ActiveSheet.Cells.Locked = True: ProtectSheet ThisWb.ActiveSheet
    
    MsgBox "Сохраните/Перезагрузите рабочую книгу " & String(2, vbCr) _
      & "Примечание: Возможное решение проблемы смотрите на странице: " _
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

Private Sub ShowCosts(ByRef control As IRibbonControl) ' rev.420
Dim Supplier As String, PartDate As Variant
  With ActiveSheet
    If .CodeName Like SHEETS_ALL Then ' CostUpdate -> GetCosts -> cnfRenew
      PartDate = .Cells(ActiveCell.Row, 6).FormulaR1C1
      Supplier = .Cells(ActiveCell.Row, 5): CostUpdate Supplier
      '.Cells(PartNumRow, 2).NumberFormat = "@" ' Ошибка в костыле rev.410
    Else ' Исходные данные до ФИО = 15
      If .CodeName = SUPP Or .CodeName = ARCH Then _
        SuppNumRow = ActiveCell.Row ' НАДО ЛИ? rev.410
      PartDate = MultidimArr(.Cells(SuppNumRow, 1).Resize(1, 15).Value, 1)
      Supplier = PartDate(10)
    End If: PartDate = GetDateAndCosts(.CodeName, PartDate) ' Statistics ->
    Debug.Print "Имя листа - "; .Name; " и cnfRenew - "; cnfRenew
    If IsArray(PartDate) Then ' rev.420
      Supplier = "Цены '" & Supplier & "' с " & CDate(PartDate(1)) & " " _
        & String(2, vbCr) & vbTab & "Группа 0: " & vbTab & PartDate(2) _
        & " руб. " & vbCr & vbTab & "Группа 1: " & vbTab & PartDate(3) _
        & " руб. " & vbCr & vbTab & "Группа 2: " & vbTab & PartDate(4) _
        & " руб. " & String(2, vbCr) & vbTab & "НАШ 1: " & String(2, vbTab) _
        & PartDate(9) & " руб. " & vbCr & vbTab & "НАШ 2: " & String(2, vbTab) _
        & PartDate(10) & " руб. " & String(2, vbCr) _
        & IIf(PartDate(6) + PartDate(7) + PartDate(8) > 0, vbTab & "НУМ 0: " _
        & String(2, vbTab) & PartDate(6) & " руб. " & vbCr & vbTab & "НУМ 1: " _
        & String(2, vbTab) & PartDate(7) & " руб. " & vbCr & vbTab & "НУМ 2: " _
        & String(2, vbTab) & PartDate(8) & " руб. " & String(2, vbCr), "")
      If UBound(PartDate) >= 14 Then Supplier = Supplier _
        & IIf(PartDate(14) > 0, vbTab & "Бухонл., Кодекс: " & vbTab & _
        PartDate(14) & " руб. " & String(2, vbCr), "") ' Ф/Л
      Supplier = Supplier & IIf(PartDate(5) > 0, vbTab & "Актуализация: " _
        & vbTab & PartDate(5) & " руб. " & String(2, vbCr), "") & vbTab _
        & "Покупка вопроса: " & vbTab & PartDate(11) & " руб. " _
        & String(2, vbCr) & vbTab & "ЮВО: " & String(2, vbTab) & PartDate(12) _
        & " руб. " & String(2, vbCr)
      If UBound(PartDate) >= 13 Then Supplier = Supplier _
        & IIf(PartDate(13) > 0, vbTab & "Официал. письма: " & vbTab _
        & PartDate(13) & " руб. " & vbCr, "") ' Ф/Л
      MsgBox Supplier, vbOKOnly, "Категория цены: " & cnfRenew
    End If
  End With
End Sub
