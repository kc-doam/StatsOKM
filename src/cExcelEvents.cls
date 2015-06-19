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
'12345678901234567890123456789012345678901234567890123456789012345678901234567890

Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1
Private Count As Long  ' Счётчик; Если с минусом, то номер строки

Private Sub Class_Initialize() ' Open #1.2
  Set App = Application
  ' Запретить ввод без перемещения курсора. Не работает при выборе из списка
  App.OnKey "^~", "SendKeyEnter": App.OnKey "^{ENTER}", "SendKeyEnter"
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook) ' Open #1.3
  ' Действий не требуется
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook) ' Open #1.4
  On Error Resume Next
    If App_Wb.FullName = Wb.FullName Then _
      App.CellDragAndDrop = False: App.MoveAfterReturnDirection = xlToRight
    If Err Then ErrCollection Err.Number, 2, 16: End ' EPN = 2
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, _
Cancel As Boolean) ' Close #1.1
  On Error Resume Next
    If App_Wb.FullName = Wb.FullName Then
      Cancel = ChangedBeforeSave(Wb.ActiveSheet) ' Отменить закрытие файла
      If Not Cancel And IsArray(SuppDiff) And SuppNumRow > 1 Then _
        Set App = Nothing: RecordCells True ' ЗАПИСЬ
      App_SheetActivate Wb.ActiveSheet ' Костыль
    End If: If Err Then ErrCollection Err.Number, 2, 16: End ' EPN = 2
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
    End If: If Err Then ErrCollection Err.Number, 2, 16: End ' EPN = 2
End Sub

Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook) ' Close #1.3
  On Error Resume Next ' Деактивация книги работает только в окне документа
    If App_Wb.FullName = Wb.FullName Then
      ''Stop
      App.CellDragAndDrop = True: App.MoveAfterReturnDirection = xlDown
      'If GetSheetList(Set_spName) > 0 Then
      '  Debug.Print Set_spName; " "; Wb.Sheets(Sh_List(Set_spName)).CodeName
      '  Debug.Print Set_spName: Stop
      If Wb.Sheets(GetSheetList(Set_spName)).Name = cnfRenew _
      And IsArray(SuppDiff) And SuppNumRow > 1 Then
        RecordCells CheckSupplier
        SuppDiff = Empty: ProtectSheet Wb.ActiveSheet
      End If
    End If: If Err Then ErrCollection Err.Number, 2, 16: End ' EPN = 2
End Sub

Private Sub App_SheetDeactivate(ByVal Sh As Object) ' #S Серый лист
  App.Cursor = xlNorthwestArrow ' Курсор «Стрелка»
  'App.ScreenUpdating = False  ' ВЫКЛ Обновление экрана
  If Count < 0 And Not Sh.ProtectScenarios Then _
    Sh.Rows(Abs(Count)).Locked = True: Count = 0 ' Защита ячеек
  If Sh.CodeName = Set_arName Then
    On Error Resume Next
      App.ScreenUpdating = False: SortSupplier Sh, 10, 11
      Sh.Visible = xlSheetHidden: App.ScreenUpdating = True ' СКРЫТЬ
      If Err Then ErrCollection Err.Number, 3, 16, Sh.Name: Exit Sub ' EPN = 3
    On Error GoTo 0
  End If
  ''Stop: Debug.Print SuppNumRow; Selection.Row
  If Sh.CodeName = Set_spName And IsArray(SuppDiff) _
  And SuppNumRow > 1 Then RecordCells CheckSupplier
  SuppDiff = Empty: ProtectSheet Sh
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object) ' #S Белый лист
  If Sh.CodeName <> Set_spName And Sh.CodeName <> Set_arName Then PartNumRow = ActiveCell.Row
  If Sh.CodeName = Set_spName And IsEmpty(SuppDiff) Then
    SuppNumRow = Selection.Row
    ' Создаём массив с изменениями о поставщике rev.310
    SuppDiff = MultidimArr(Sh.Cells(SuppNumRow, 1) _
      .Resize(, 15).Value, 1) ' Исходные данные до ФИО = 15
  End If: cnfRenew = Sh.Name ': App.ScreenUpdating = True ' ВКЛ Обновление экрана
  Application.Cursor = xlDefault ' Восстановить Курсор «по умолчанию»
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
  If (Sh.CodeName = "SF_" Or Sh.CodeName = "SB_") And Target.Locked Then _
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
        Or IsEmpty(Sh.Cells(Target.Row, 11)))) _
      Or (Sh.CodeName = "SF_" And IsEmpty(Sh.Cells(Target.Row, 18))) _
      Or (Sh.CodeName = "SB_" And IsEmpty(Sh.Cells(Target.Row, 18))) Then _
        Sh.Rows(Target.Row).Locked = False ' Снять защиту ячеек
    End If
  End If
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
Dim cell As Range ' Не работает при выборе из списка
  If Count < 0 Then Sh.Rows(Abs(Count)).Locked = True: Count = 0 ' Защита ячеек
  If Sh.CodeName = Set_spName And Target.Column = 11 Then Exit Sub
  
  'Application.ScreenUpdating = False ' ВЫКЛ Обновление экрана
  If Sh.CodeName = Set_spName And SuppNumRow <> Target.Row Then
    If IsArray(SuppDiff) And SuppNumRow > 1 Then RecordCells CheckSupplier
    ' Создаём массив с изменениями о поставщике rev.310
    If Not SuppNumRow = Selection.Row Then SuppDiff = Empty
    App_SheetActivate Sh ' Костыль
  End If
  If Sh.CodeName = Set_spName And Target.Row > 1 Then
    If Target.Column = 12 And Not IsEmpty(Sh.Cells(Target.Row, 10)) Then
      UnprotectSheet Sh
      ListCost Sh, Target.Row
      ProtectSheet Sh
    End If
    'Debug.Print SuppNumRow; " "; Selection.Row
    '' Создаём массив с изменениями о поставщике rev.310
    'If Not SuppNumRow = Selection.Row Then SuppDiff = Empty
    'App_SheetActivate Sh ' Костыль
  End If
  
  If (Sh.CodeName = "SF_" Or Sh.CodeName = "SB_") And Target.Row > 1 Then
    UnprotectSheet Sh
    If Target.Column = 5 And Not IsEmpty(Sh.Cells(Target.Row - 1, 5)) Then
      With Sh.Cells(Target.Row, 5).Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
          Formula1:="=" & Let_SuppList ' Добавляем список в ячейку
        .ErrorTitle = "Поставщик (кратко)"
        .ErrorMessage = "Необходимо выбрать значение из списка "
        .ShowError = True: .IgnoreBlank = True
      End With
    End If
    If (Target.Column = 6 Or Target.Column = 7) And PartNumRow > 1 Then
      With Sh.Cells(PartNumRow, Target.Column)
        'Stop ' !!!
        GetSuppRow Sh, .Row ' Поиск строки SuppNumRow у поставщика
        If Not IsEmpty(Sh.Cells(.Row, .Column - 1)) And SuppNumRow > 0 Then
          'cnfRenew = Sh.Parent.Sheets(Sh_List(Set_arName)).Name
          Sh.Cells(.Row, 1).FormulaR1C1 = "=TEXT(R" & .Row & "C9,""ММММ.ГГ"")"
          Sh.Cells(.Row, 2).FormulaR1C1 = "='" & cnfRenew & "'!R" & SuppNumRow & "C4"
          Sh.Cells(.Row, 2).NumberFormat = "General"
          Sh.Cells(.Row, 3).FormulaR1C1 = "='" & cnfRenew & "'!R" & SuppNumRow & "C5"
          Sh.Cells(.Row, 4).FormulaR1C1 = "='" & cnfRenew & "'!R" & SuppNumRow & "C8"
          Sh.Cells(.Row, 6).NumberFormat = "dd/mm/yyyy"
          ''Stop ' STOP
          Set cell = Sh.Parent.Sheets(Sh_List(Set_arName)).Cells(SuppNumRow, 1)
          ' ЦЕНЫ КФ
          Sh.Cells(.Row, 26).FormulaR1C1 = "=SUM(R" & .Row & "C23:R" & .Row & "C25)"
          Sh.Cells(.Row, 30).FormulaR1C1 = "=SUM(R" & .Row & "C27:R" & .Row & "C29)"
          Sh.Cells(.Row, 40).FormulaR1C1 = "=SUM(R" & .Row & "C36:R" & .Row & "C39)"
          Sh.Cells(.Row, 41).FormulaR1C1 = GetCosts(Sh, .Row, cell)
          Sh.Cells(.Row, 42).FormulaR1C1 = "=IF('" & cnfRenew & "'!R" & SuppNumRow & "C13=""НДС"",R" & .Row & "C41*0.18,IF('" & cnfRenew & "'!R" & SuppNumRow & "C13=""УСН"",""без НДС"",""""))"
          Sh.Cells(.Row, 44).FormulaR1C1 = "=SUM(R" & .Row & "C41:R" & .Row & "C43)"
          ' ЦЕНЫ БО
          'Sh.Cells(.Row, 57).FormulaR1C1 = "=SUM(R" & .Row & "C53:R" & .Row & "C56)"
          'Sh.Cells(.Row, 70).FormulaR1C1 = GetCosts(Sh, .Row, cell)
          'Sh.Cells(.Row, 71).FormulaR1C1 = "=IF('" & cnfRenew & "'!R" & SuppNumRow & "C13=""НДС"",R" & .Row & "C70*0.18,IF('" & cnfRenew & "'!R" & SuppNumRow & "C13=""УСН"",""без НДС"",""""))"
          'Sh.Cells(.Row, 73).FormulaR1C1 = "=SUM(R" & .Row & "C70:R" & .Row & "C72)"
          'Sh.Cells(.Row, 74).FormulaR1C1 = "=R" & .Row & "C64+R" & .Row & "C73"
        Else
          For Each cell In Sh.Range(Sh.Cells(.Row, 1), Sh.Cells(.Row, 74))
            If cell.HasFormula Then cell = ""
          Next cell
        End If
      End With: PartNumRow = Target.Row
    End If: ProtectSheet Sh
  End If ': Application.ScreenUpdating = True ' ВКЛ Обновление экрана
End Sub