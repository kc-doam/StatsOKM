VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cExcelEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890

Private WithEvents App As Application, Clipbord_RangeAddress As String
Attribute App.VB_VarHelpID = -1
Private Count As Long  ' Счётчик; Если с минусом, то номер строки

Private Sub Class_Initialize() ' Open #1.2
  Set App = Application
  ' Запретить ввод без перемещения курсора. Не работает при выборе из списка
  App.OnKey "^~", "SendKeyEnter": App.OnKey "^{ENTER}", "SendKeyEnter"
  ' Вставка данных через горячие клавиши rev.330
  App.OnKey "^v", "SendKeysCtrlV": App.OnKey "+{INSERT}", "SendKeysCtrlV"
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook) ' Open #1.3
  ' Действий не требуется
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook) ' Open #1.4
  On Error Resume Next
    If App_Wb.FullName = Wb.FullName Then Quit = False
    If Err Then ErrCollection Err.Number, 2, 16: Quit = 1 ' EPN = 2
    Clipbord_RangeAddress = Selection.address ' Для буфера
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, _
Cancel As Boolean) ' Close #1.1
  On Error Resume Next
    If App_Wb.FullName = Wb.FullName Then
      Cancel = ChangedBeforeSave(Wb.ActiveSheet) ' Отменить закрытие файла
      If Not Cancel And IsArray(SuppDiff) And SuppNumRow > 1 Then _
        Set App = Nothing: RecordCells True ' ЗАПИСЬ
      App_SheetActivate Wb.ActiveSheet ' Костыль
    End If: If Err Then ErrCollection Err.Number, 2, 16: Quit = 1 ' EPN = 2
End Sub

Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, _
ByVal SaveAsUI As Boolean, Cancel As Boolean) ' Close #1.2
  On Error Resume Next ' SaveAsUI не работает, причина не известна
    If App_Wb.FullName = Wb.FullName Then
      Cancel = ChangedBeforeSave(Wb.ActiveSheet) ' Отменить сохранения файла
      'Debug.Print Cancel; Len(cnfRenew): Stop
      If Not Cancel And Len(cnfRenew) > 0 Then
        If IsArray(SuppDiff) And SuppNumRow > 1 Then RecordCells True ' ЗАПИСЬ
        SpecificationSheets Wb.ActiveSheet.Index
      End If
    End If: If Err Then ErrCollection Err.Number, 2, 16: Quit = 1 ' EPN = 2
End Sub

Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook) ' Close #1.3
  ' Отключено в rev.330
  On Error Resume Next ' Деактивация книги работает только в окне документа
    If App_Wb.FullName = Wb.FullName Then
      App.CellDragAndDrop = True: App.MoveAfterReturnDirection = xlDown
      If App.CutCopyMode Then ' Если ячейки скопированы в буфер rev.340
        App_Wb.Sheets(cnfRenew).Range(Clipbord_RangeAddress).Copy
      End If
      'If GetSheetList(Set_spName) > 0 Then
      '  Debug.Print Set_spName; " "; Wb.Sheets(Sh_List(Set_spName)).CodeName
      '  Debug.Print Set_spName: Stop
      If Wb.Sheets(GetSheetList(Set_spName)).Name = cnfRenew _
      And IsArray(SuppDiff) And SuppNumRow > 1 Then
        RecordCells CheckSupplier
        ' см.App_SheetDeactivate - SuppDiff = Empty: ProtectSheet Wb.ActiveSheet
      End If
    End If: If Err Then ErrCollection Err.Number, 2, 16: Quit = 1 ' EPN = 2
End Sub

Private Sub App_SheetDeactivate(ByVal Sh As Object) ' #S Серый лист
  App.Cursor = xlNorthwestArrow ' Курсор «Стрелка»
  'App.ScreenUpdating = False  ' ВЫКЛ Обновление экрана
  If Count < 0 And Not Sh.ProtectScenarios Then _
    Sh.Rows(Abs(Count)).Locked = True: Count = 0 ' Защита ячеек
  If Sh.CodeName = Set_arName Then
    On Error Resume Next
      App.ScreenUpdating = False: SortSupplier Sh, 10, 15 ' rev.360
      Sh.Visible = xlSheetHidden: App.ScreenUpdating = True ' СКРЫТЬ
      If Err Then ErrCollection Err.Number, 3, 16, Sh.Name: Exit Sub ' EPN = 3
    On Error GoTo 0
  End If
  ''Stop: Debug.Print SuppNumRow; Selection.Row
  If Sh.CodeName = Set_spName And IsArray(SuppDiff) _
  And SuppNumRow > 1 Then RecordCells CheckSupplier
  ' Нельзя применять ProtectSheet Sh - стирается буфер
  ' Нельзя применять SuppDiff = Empty - возникает ошибка
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object) ' #S Белый лист
  If Sh.CodeName <> Set_spName And Sh.CodeName <> Set_arName Then _
    PartNumRow = ActiveCell.Row
  If Sh.CodeName = Set_spName And IsEmpty(SuppDiff) Then
    SuppNumRow = Selection.Row
    ' Создаём массив с изменениями о поставщике rev.310
    SuppDiff = MultidimArr(Sh.Cells(SuppNumRow, 1) _
      .Resize(, 15).Value, 1) ' Исходные данные до ФИО = 15
  End If ': App.ScreenUpdating = True ' ВКЛ Обновление экрана
  Application.Cursor = xlDefault ' Восстановить Курсор «по умолчанию»
  Clipbord_RangeAddress = Selection.address ' Для буфера
  If Len(cnfRenew) > 0 Then AutoOpenControls: cnfRenew = Sh.Name ' rev.330
End Sub



Private Sub App_SheetBeforeDoubleClick(ByVal Sh As Object, _
ByVal Target As Range, Cancel As Boolean)
  App.ScreenUpdating = False ' ВЫКЛ Обновление экрана
  If Sh.CodeName = Set_spName And Not IsEmpty(Sh.Cells(Target.Row, 10)) _
  And Target.Row > 1 Then
    With UnprotectSheet(App_Wb.Sheets(Sh_List(Set_arName)))
      .Activate: .Visible = xlSheetVisible
      If Not IsEmpty(.Cells(1, 10)) Then .Range("A1:O4") _
        .AutoFilter Field:=10, Criteria1:=Sh.Cells(Target.Row, 10) _
      Else Sh.Activate: Cancel = True ' Отменяем событие «Двойной клик»
    End With: ProtectSheet App_Wb.Sheets(Sh_List(Set_arName))
  End If
  Stop
  If Sh.CodeName Like "[OQS]?_" And Target.Locked Then _
    Cancel = True: SendKeys "{F2}", False ' rev.250 Фокус должен быть на MS Excel
  App.ScreenUpdating = True ' ВКЛ Обновление экрана
End Sub

Private Sub App_SheetBeforeRightClick(ByVal Sh As Object, _
ByVal Target As Range, Cancel As Boolean)
  If Target.Worksheet.CodeName = Sh.CodeName Then ' Если лист тот же
    ' Проверка выделения
    If Not Intersect(Target, Sh.Range("A:A")) Is Nothing _
    And Sh.Cells.Columns.Count = Target.Cells.CountLarge Then
      Count = -1 * Target.Row
      ' Если «Разрешено включать» = Пусто, то разблокировать строку
      If (Sh.CodeName = Set_spName And (IsEmpty(Sh.Cells(Target.Row, 10)) _
        Or IsEmpty(Sh.Cells(Target.Row, 15)))) _
      Or (Sh.CodeName = "SF_" And IsEmpty(Sh.Cells(Target.Row, 18))) _
      Or (Sh.CodeName = "SB_" And IsEmpty(Sh.Cells(Target.Row, 18))) Then _
        Sh.Rows(Target.Row).Locked = False ' Снять защиту ячеек rev.360
    End If
  End If
End Sub

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
  If Target.Worksheet.CodeName = Sh.CodeName Then ' Если лист тот же
    If Sh.CodeName Like "[OQS]?_" And Target.Row > 1 Then
      ' 18/04/2014 Автопростановка "Разрешено включать" по оперативному договору
      If Not IsEmpty(Sh.Cells(Target.Row, 8)) And Target.Column = 19 Then
        If Sh.Cells(Target.Row, Target.Column) = "опер." Then _
          Sh.Cells(Target.Row, 20) = Sh.Cells(Target.Row, 12) _
        Else Sh.Cells(Target.Row, 20).ClearContents
      End If
      ' 08/12/2014 Удаление всех символов, кроме цифр
      If Target.Column > 24 And Target.Column < 44 Then
        If Not IsArray(Target) And Not Target.HasFormula Then ' Если не массив
          Count = Replace(Val(Target.Value), "-", "")
          If Not Target = Count Then Target = Count
        End If
      End If
      If App.CountA(Sh.Range(Sh.Cells(Target.Row, 29), _
        Sh.Cells(Target.Row, 31))) > 2 And Sh.Cells(Target.Row, 32) < 1 Then _
        If IsEmpty(Sh.Cells(Target.Row, 24)) Then Sh.Cells(Target.Row, 24) = _
        "не оплач." ' Автопростановка "не оплач." rev.360
    End If
  End If
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
Dim cell As Range ' Не работает при выборе из списка
  If Count < 0 Then Sh.Rows(Abs(Count)).Locked = True: Count = 0 ' Защита ячеек
  If Sh.CodeName = Set_spName And Target.Column = 15 Then Exit Sub ' rev.360
  
  'Application.ScreenUpdating = False ' ВЫКЛ Обновление экрана
  If Sh.CodeName = Set_spName And SuppNumRow <> Target.Row Then
    If IsArray(SuppDiff) And SuppNumRow > 1 Then RecordCells CheckSupplier
    ' Создаём массив с изменениями о поставщике rev.310
    If Not SuppNumRow = Selection.Row Then SuppDiff = Empty
    App_SheetActivate Sh ' Костыль
  End If
  If Sh.CodeName = Set_spName And Target.Row > 1 Then
    If Target.Column = 11 And Not IsEmpty(Sh.Cells(Target.Row, 10)) Then ' rev.360
      UnprotectSheet Sh
      ListCost Sh, Target.Row
      ProtectSheet Sh
    End If
    'Debug.Print SuppNumRow; " "; Selection.Row
    '' Создаём массив с изменениями о поставщике rev.310
    'If Not SuppNumRow = Selection.Row Then SuppDiff = Empty
    'App_SheetActivate Sh ' Костыль
  End If
  
  If Sh.CodeName Like "[OQS]?_" And Target.Row > 1 Then
    If (Target.Column = 5 Or Target.Column = 19 Or Target.Column = 48) _
    And Not IsEmpty(Sh.Cells(Target.Row - 1, 5)) Then
      ' And IsEmpty(Sh.Cells(Target.Row, 5)) ' ОШИБКА! В прошлых записях нет списка rev.330
      UnprotectSheet Sh ' Форма договора rev.310
      With Sh.Cells(Target.Row, 5).Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
          Formula1:="=" & Let_SuppList ' Добавляем список в ячейку
        .ErrorTitle = "Поставщик (кратко)"
        .ErrorMessage = "Необходимо выбрать значение из списка "
        .ShowError = True: .IgnoreBlank = True
      End With
      With Sh.Cells(Target.Row, 19).Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
          Formula1:=Let_ContrFormList
        .ErrorTitle = "Форма договора"
        .ErrorMessage = "Необходимо выбрать значение из списка "
        .IgnoreBlank = True: .InCellDropdown = True
        .Parent.Locked = False
      End With
      With Sh.Cells(Target.Row, 48).Validation ' rev.350
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
          Formula1:=Replace(Let_UrgencyTransList, "$", "Срочно ")
        .ErrorTitle = "Срочно"
        .ErrorMessage = "Необходимо выбрать значение из списка "
        .IgnoreBlank = True: .InCellDropdown = True
        .Parent.Locked = False
      End With: ProtectSheet Sh ' rev.330
    End If
    If (Target.Column = 6 Or Target.Column = 7) And PartNumRow > 1 Then
      With Sh.Cells(PartNumRow, Target.Column)
        'Stop ' !!!
        Debug.Print "Sh.Row ="; .Row; " Sh.Column ="; .Column
        GetSuppRow Sh, .Row ' Поиск строки SuppNumRow у поставщика
        If Not IsEmpty(Sh.Cells(.Row, 5)) _
        Or Not IsEmpty(Sh.Cells(.Row, .Column - 1)) Then ' rev.340 And SuppNumRow > 0
          SetFormula Sh, .Row ' Создание формул для записи rev.350
        Else
          For Each cell In Sh.Range(Sh.Cells(.Row, 1), Sh.Cells(.Row, 74))
            If cell.HasFormula Then cell = ""
          Next cell
        End If
      End With: PartNumRow = Target.Row
    Else: CostChanged ' Необходимость обновить статистику rev.340
    End If
  End If: AutoOpenControls ' rev.330
  ': Application.ScreenUpdating = True ' ВКЛ Обновление экрана
  Clipbord_RangeAddress = Selection.address ' Для буфера
End Sub
