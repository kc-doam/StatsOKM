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

Private WithEvents App As Application, Clipboard_RangeAddress As String
Attribute App.VB_VarHelpID = -1
Private Count As Long ' Счётчик; Если с минусом, то номер строки

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
    If ThisWb.FullName = Wb.FullName Then ' rev.370
      Quit = False: Clipboard_RangeAddress = Selection.Address ' Для буфера
      ' Удалить меню: Очистить содержимое (Работает на Защищённом листе) rev.410
      App.CommandBars("Row").FindControl(Id:=3125).Delete
      With App.CommandBars("Cell") ' Удаление меню специальной вставки
        ' http://www.rondebruin.nl/win/s4/win011.htm
        With .Controls.Add(Id:=370, Temporary:=True, _
          Before:=.Controls(.FindControl(Id:=22).Caption).Index)
          .Caption = "&Вставить значения": .FaceId = 22 ' Недокументировано
        End With
        .FindControl(Id:=22).Delete ' Удалить меню: Вставить
        .FindControl(Id:=3624).Delete ' Удалить меню: Вставить таблицу
      End With: App.Calculation = xlAutomatic ' rev.370
    End If: If Err Then ErrCollection Err.Number, 2, 16: Quit = 1 ' EPN = 2
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, _
Cancel As Boolean) ' Close #1.1
  On Error Resume Next
    If ThisWb.FullName = Wb.FullName Then
      Cancel = ChangedBeforeSave(Wb.ActiveSheet) ' Отменить закрытие файла
      If Not Cancel And IsArray(SuppDiff) And SuppNumRow > 1 Then _
        RecordCells True ' ЗАПИСЬ, Удалено 'Set App = Nothing' rev.390
      App_SheetActivate Wb.ActiveSheet ' Костыль
    End If: If Err Then ErrCollection Err.Number, 2, 16: Quit = 1 ' EPN = 2
End Sub

Private Sub App_WorkbookDeactivate(ByVal Wb As Workbook) ' Close #1.2
Dim CopyMode As Boolean: CopyMode = App.CutCopyMode ' Буфер заполнен rev.370
  ' Отключено в rev.330
  On Error Resume Next ' Деактивация книги работает только в окне документа
    If ThisWb.FullName = Wb.FullName Then
      App.CellDragAndDrop = True: App.MoveAfterReturnDirection = xlDown
      App.DisplayPasteOptions = True: App.CommandBars("Cell").Reset ' rev.370
      App.CommandBars("Row").Reset ' rev.410
      If CopyMode Then _
        ThisWb.Sheets(cnfRenew).Range(Clipboard_RangeAddress).Copy ' rev.370
      'If GetSheetIndex(SUPP) > 0 Then
      '  Debug.Print SUPP; " "; Wb.Sheets(Sh_List(SUPP)).CodeName
      'If ActiveSheet.CodeName = SUPP Then Stop
      If Wb.Sheets(GetSheetIndex(SUPP)).Name = cnfRenew _
      And IsArray(SuppDiff) And SuppNumRow > 1 Then
        RecordCells CheckSupplier
        ' см.App_SheetDeactivate - SuppDiff = Empty: ProtectSheet Wb.ActiveSheet
      End If
    End If: If Err Then ErrCollection Err.Number, 2, 16: Quit = 1 ' EPN = 2
End Sub

Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, _
ByVal SaveAsUI As Boolean, Cancel As Boolean) ' Close #1.3
  On Error Resume Next ' SaveAsUI не работает, причина не известна
    If ThisWb.FullName = Wb.FullName Then
      Cancel = ChangedBeforeSave(Wb.ActiveSheet) ' Отменить сохранения файла
      'Debug.Print Cancel; Len(cnfRenew): Stop
      If Not Cancel And Len(cnfRenew) > 0 Then
        If IsArray(SuppDiff) And SuppNumRow > 1 Then RecordCells True ' ЗАПИСЬ
        Count = 0: SpecificationSheets Wb.ActiveSheet.Index ' rev.380
      End If
    End If: If Err Then ErrCollection Err.Number, 2, 16: Quit = 1 ' EPN = 2
End Sub

Private Sub App_SheetDeactivate(ByVal Sh As Object) ' #S «Серый лист»
  App.Cursor = xlNorthwestArrow ' Курсор «Стрелка»
  'App.ScreenUpdating = False  ' ВЫКЛ Обновление экрана
  If Count < 0 And Not Sh.ProtectScenarios Then _
    Sh.Rows(Abs(Count)).Locked = True: Count = 0 ' Защита ячеек
  If Sh.CodeName = ARCH Then
    On Error Resume Next
      App.ScreenUpdating = False: SortSupplier Sh, 10, 15 ' rev.360
      Sh.Visible = xlSheetHidden: App.ScreenUpdating = True ' СКРЫТЬ
      If Err Then ErrCollection Err.Number, 3, 16, Sh.Name: Exit Sub ' EPN = 3
    On Error GoTo 0
  End If
  ''Stop: Debug.Print SuppNumRow; Selection.Row
  If Sh.CodeName = SUPP And IsArray(SuppDiff) _
  And SuppNumRow > 1 Then RecordCells CheckSupplier
  ' Нельзя применять ProtectSheet Sh - стирается буфер
  ' Нельзя применять SuppDiff = Empty - возникает ошибка
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object) ' #S «Белый лист»
  With ActiveSheet
    If .CodeName Like SHEETS_ALL Or .CodeName = ARCH Then _
      If CostChanged Then CostUpdate
      'SuppDiff = Empty ' !!! Передаётся в Auto_Open -> Settings rev.410
    If .CodeName = SUPP And Not IsArray(SuppDiff) Then ' rev.410
      Erase Get_Supp: SuppNumRow = Selection.Row ' rev.390
      ' ВАЖНО! Может создаться массив с пустыми значениями
      ' Создаём массив с изменениями о поставщике (до ФИО = 15) rev.310
      SuppDiff = MultidimArr(.Cells(SuppNumRow, 1).Resize(1, 15).Value, 1)
      If Len(SuppDiff(15)) < &H4 Then SuppDiff(15) = Settings("date0") ' rev.420
    ElseIf .CodeName Like SHEETS_ALL Then ' rev.420
      PartNumRow = ActiveCell.Row
    End If: AutoOpenControls ': App.ScreenUpdating = True ' ВКЛ Обновление экрана
    cnfRenew = .Name: App.Cursor = xlDefault ' Восстановить Курсор «по умолчанию»
    'Clipboard_RangeAddress = Selection.address ' Для буфера - откл rev.370
  End With
End Sub



Private Sub App_SheetBeforeDoubleClick(ByVal Sh As Object, _
ByVal Target As Range, Cancel As Boolean)
  'App.ScreenUpdating = False ' ВЫКЛ Обновление экрана
  If Sh.CodeName = SUPP And Not IsEmpty(Sh.Cells(Target.Row, 10)) _
  And Target.Row > 1 And Target.Column <= 15 Then ' rev.400
    With UnprotectSheet(ThisWb.Sheets(Sh_List(ARCH)))
      .Activate: .Visible = xlSheetVisible
      If ActiveSheet.CodeName = .CodeName And Not IsEmpty(.Cells(1, 10)) Then
        .ShowAllData: .Cells(1, 10).Select ' rev.420
        .Range("A1:O4").AutoFilter Field:=10, Criteria1:=Sh.Cells(Target.Row, 10)
        ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1 ' CTRL+HOME
      Else
        Sh.Activate: Cancel = True ' Отменяем событие "Двойной клик"
      End If
    End With: ProtectSheet ThisWb.Sheets(Sh_List(ARCH))
  End If
  If Sh.CodeName Like SHEETS_ALL And Target.Locked Then _
    Cancel = True: SendKeys "{F2}", False ' rev.250 Фокус должен быть на MS Excel
    'If Target.AllowEdit Then MsgBox ' ЗАМЕНИТЬ! rev.370
  'App.ScreenUpdating = True ' ВКЛ Обновление экрана
End Sub

Private Sub App_SheetBeforeRightClick(ByVal Sh As Object, _
ByVal Target As Range, Cancel As Boolean)
  If Target.Worksheet.CodeName = Sh.CodeName Then ' Если лист тот же
    ' Проверка выделения
    If Not Intersect(Target, Sh.Range("A:A")) Is Nothing _
    And Sh.Cells.Columns.Count = Target.Cells.CountLarge Then
      Count = -1 * Target.Row
      ' Если «Разрешено включать» = Пусто, то разблокировать строку
      If (Sh.CodeName = SUPP And (IsEmpty(Sh.Cells(Target.Row, 10)) _
        Or IsEmpty(Sh.Cells(Target.Row, 15)))) _
      Or (Sh.CodeName = "SF_" And IsEmpty(Sh.Cells(Target.Row, 18))) _
      Or (Sh.CodeName = "SB_" And IsEmpty(Sh.Cells(Target.Row, 18))) Then _
        Sh.Rows(Target.Row).Locked = False ' Снять защиту ячеек rev.360
    End If
  End If
End Sub

' Процедура App_SheetChange может выполняться 2-й раз
Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
'  Debug.Print "+ Row ="; Target.Row; " Column ="; Target.Column
  If Target.Worksheet.CodeName = Sh.CodeName Then ' Если лист тот же
    If Sh.CodeName Like "S[BF]_" And Target.Row > 1 Then ' rev.410
      ' 18/04/2014 Автопростановка "Разрешено включать" по оперативному договору
      If Not IsEmpty(Sh.Cells(Target.Row, 8)) And Target.Column = 19 Then
        If Sh.Cells(Target.Row, Target.Column) = "опер." Then _
          Sh.Cells(Target.Row, 20) = Sh.Cells(Target.Row, 12) _
        Else Sh.Cells(Target.Row, 20).ClearContents
      End If
      ' 08/12/2014 Удаление всех символов, кроме цифр rev.390
      If Target.Column > 24 And Target.Column < 46 Then ' rev.410
        If Not IsArray(Target) And Not Target.HasFormula Then ' Если не массив
          Count = Replace(Val(Target.Value), "-", "")
          If Not Target = Count Then Target = Count
        End If
      End If
      If App.CountA(Sh.Range(Sh.Cells(Target.Row, 31), _
        Sh.Cells(Target.Row, 33))) > 2 And Sh.Cells(Target.Row, 34) < 1 Then _
        If IsEmpty(Sh.Cells(Target.Row, 24)) Then Sh.Cells(Target.Row, 24) = _
        "не оплач." ' Автопростановка "не оплач." rev.390
    End If
    If Sh.CodeName = SUPP Then ' rev.420
      App.ScreenUpdating = False: App.EnableEvents = False ' ВЫКЛ События
      If IsArray(Target.Value) Then
        For Count = LBound(Target.Value) To UBound(Target.Value)
          If Count = LBound(Target.Value) Then Target(Count) = _
            ClearSpacesInText(Target(Count)) Else Target(Count) = Get_Supp( _
            SuppNumRow + Count - 1) ' Костыль (1 = должен быть Settings("head"))
        Next Count: ErrCollection 30, 1, 64 ' EPN = 1
      Else: Target = ClearSpacesInText(Target): End If
      Target.WrapText = False ' ВЫКЛ Перенос по словам
      App.ScreenUpdating = True: App.EnableEvents = True ' ВКЛ События
    End If
  End If
  If ActiveSheet.CodeName Like "S[BF]_" And Target.Column = 29 Then ' rev.420
    With ThisWb.Sheets(Sh_List(IIf(Sh.CodeName = "SF_", "SB_", "SF_")))
      Clipboard_RangeAddress = "$E$" & .Cells.SpecialCells(xlLastCell).Row + 1
      If .Range(Clipboard_RangeAddress).AllowEdit And WorksheetFunction.Sum( _
        Target.Offset(, -3).Resize(1, 4)) = Target.Offset(, -4) _
      And Target > 0 Then
        If MsgBox("Указано '" & Target & "' переходящих материала в " _
          & IIf(.CodeName = "SF_", "Финансист", "Бюджет") & ". " & vbCrLf _
          & "Добавить новую партию на лист '" & .Name & "'? ", 260 + 48, _
          "Добавление партии") = vbYes Then ' rev.410
          .Cells(.Range(Clipboard_RangeAddress).Row, 51) = "Переходящие из " _
            & IIf(.CodeName = "SF_", "БО", "КФ") ' rev.410
          Sh.Range("AX" & PartNumRow).Copy _
            Destination:=.Cells(.Range(Clipboard_RangeAddress).Row, 50)
          Sh.Range("AC" & PartNumRow).Copy _
            Destination:=.Cells(.Range(Clipboard_RangeAddress).Row, 25)
          .Range(Clipboard_RangeAddress).Resize(1, 3) = _
            Array(Sh.Range("E" & PartNumRow), Sh.Range("K" & PartNumRow), _
            Sh.Range("G" & PartNumRow)) ' из ОКМ в Дату поступления rev.410
          Clipboard_RangeAddress = "$E$" & .Cells.SpecialCells(xlLastCell).Row
          .Activate: .Range(Clipboard_RangeAddress).Select ' Костыль
          .Range(Clipboard_RangeAddress).Offset(, 1).Select ' Костыль rev.420
          .Cells(PartNumRow, 2).NumberFormat = "@" ' Костыль
        End If
      End If: If Not .Range(Clipboard_RangeAddress).AllowEdit Then _
        ErrCollection 5, 3, 16, Clipboard_RangeAddress _
        & " на листе '" & .Name ' EPN = 3
    End With
  End If
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
'Dim cell As Range ' Не работает при выборе из списка
  If Count < 0 Then Sh.Rows(Abs(Count)).Locked = True: Count = 0 ' Защита ячеек
  If Sh.CodeName = SUPP And Target.Row > 1 Then ' Перемещено rev.420
    If Target.Column = 11 And Not IsEmpty(Sh.Cells(Target.Row, 10)) Then _
      UnprotectSheet Sh: ListCost Sh, Target.Row: ProtectSheet Sh
    'Debug.Print SuppNumRow; " "; Selection.Row
    '' Создаём массив с изменениями о поставщике rev.310
    'If Not SuppNumRow = Selection.Row Then SuppDiff = Empty
    'App_SheetActivate Sh ' Костыль
  End If
  'If Sh.CodeName = SUPP And Target.Column = 15 Then Exit Sub ' rev.360
  If Sh.CodeName = SUPP And SuppNumRow = Target.Row Then Exit Sub ' rev.410
  'Application.ScreenUpdating = False ' ВЫКЛ Обновление экрана
  
  If Sh.CodeName = SUPP Then ' And SuppNumRow <> Target.Row
    If IsArray(SuppDiff) And SuppNumRow > 1 Then RecordCells CheckSupplier
    ' Создаём массив с изменениями о поставщике rev.310
    If Not SuppNumRow = Selection.Row Then SuppDiff = Empty
    App_SheetActivate Sh ' Костыль
  End If
  
  If Sh.CodeName Like "Q?_" And Target.Row > 1 Then ' rev.410
    If (Target.Column = 5 Or Target.Column = 16 Or Target.Column = 31) _
    And Not IsEmpty(Sh.Cells(Target.Row - 1, 5)) Then
      ' And IsEmpty(Sh.Cells(Target.Row, 5)) ' ОШИБКА! В прошлых записях нет списка rev.330
      UnprotectSheet Sh ' Форма договора
      With Sh.Cells(Target.Row, 5).Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
          Formula1:="=" & SUPP_LIST ' Добавляем список в ячейку
        .ErrorTitle = "Поставщик (кратко)"
        .ErrorMessage = "Необходимо выбрать значение из списка "
        .ShowError = True: .IgnoreBlank = True
      End With
      With Sh.Cells(Target.Row, 16).Validation ' rev.410
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
          Formula1:=CONTROFORM_LIST
        .ErrorTitle = "Форма договора"
        .ErrorMessage = "Необходимо выбрать значение из списка "
        .IgnoreBlank = True: .InCellDropdown = True
        .Parent.Locked = False
      End With
      If Sh.CodeName Like "QB_" Then ' rev.410
        With Sh.Cells(Target.Row, 31).Validation
          .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:=PARTNEROFORM_LIST
          .ErrorTitle = "Форма сотрудничества"
          .ErrorMessage = "Необходимо выбрать значение из списка "
          .IgnoreBlank = True: .InCellDropdown = True
          .Parent.Locked = False
        End With
      End If: ProtectSheet Sh
    End If
  End If
  If Sh.CodeName Like "OE_" And Target.Row > 1 Then ' rev.410
    If (Target.Column = 5) And Not IsEmpty(Sh.Cells(Target.Row - 1, 5)) Then
      ' And IsEmpty(Sh.Cells(Target.Row, 5)) ' ОШИБКА! В прошлых записях нет списка rev.330
      UnprotectSheet Sh ' Форма договора
      With Sh.Cells(Target.Row, 5).Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
          Formula1:="=" & SUPP_LIST ' Добавляем список в ячейку
        .ErrorTitle = "Поставщик (кратко)"
        .ErrorMessage = "Необходимо выбрать значение из списка "
        .ShowError = True: .IgnoreBlank = True
      End With: ProtectSheet Sh
    End If
  End If
  
  ' ВАЖНО!
  If Target.Row > 1 And Sh.CodeName = "OE_" Or Sh.CodeName Like "Q?_" Then
    PartNumRow = Target.Row ' rev.380
    If Target.Column = 6 Or Target.Column = IIf(Sh.CodeName Like "Q?_", 7, 8) _
    And PartNumRow > 1 Then
      App.ScreenUpdating = False ' ВЫКЛ Обновление экрана
      If Not IsEmpty(Sh.Cells(Target.Row, 5)) Then _
        CostUpdate Sh.Cells(Target.Row, 5) ' Костыль And SuppNumRow > 0
    End If: cnfRenew = Sh.Name ' rev.410
  End If
  'If (Sh.CodeName Like "S?_" Or Sh.CodeName Like "O?_") _
  And Target.Row > 1 Then ' rev.410
  If Target.Row > 1 And (Sh.CodeName Like SHEETS_ALL Xor Sh.CodeName = "OE_" _
  Xor Sh.CodeName = "QB_" Xor Sh.CodeName = "QT_") Then ' rev.410
    If (Target.Column = 5 Or Target.Column = 17 Or Target.Column = 19 _
    Or Target.Column = 31 Or Target.Column = 50) _
    And Not IsEmpty(Sh.Cells(Target.Row - 1, 5)) Then ' rev.410
      ' And IsEmpty(Sh.Cells(Target.Row, 5)) ' ОШИБКА! В прошлых записях нет списка rev.330
      UnprotectSheet Sh ' Форма договора rev.310
      With Sh.Cells(Target.Row, 5).Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
          Formula1:="=" & SUPP_LIST ' Добавляем список в ячейку
        .ErrorTitle = "Поставщик (кратко)"
        .ErrorMessage = "Необходимо выбрать значение из списка "
        .ShowError = True: .IgnoreBlank = True
      End With
      With Sh.Cells(Target.Row, IIf(Sh.CodeName Like "SL_", 17, 19)) _
        .Validation ' rev.410
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
          Formula1:=CONTROFORM_LIST
        .ErrorTitle = "Форма договора"
        .ErrorMessage = "Необходимо выбрать значение из списка "
        .IgnoreBlank = True: .InCellDropdown = True
        .Parent.Locked = False
      End With
      If Sh.CodeName Like "OU_" Then ' rev.410
        With Sh.Cells(Target.Row, 31).Validation
          .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:=PARTYTYPE_LIST
          .ErrorTitle = "Тип партии"
          .ErrorMessage = "Необходимо выбрать значение из списка "
          .IgnoreBlank = True: .InCellDropdown = True
          .Parent.Locked = False
        End With
      End If
      If Sh.CodeName Like "S[BF]_" Then ' rev.410
        With Sh.Cells(Target.Row, 50).Validation ' rev.390
          .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
            Formula1:=Replace(EXTRA_LIST, "$", "Срочно ")
          .ErrorTitle = "Срочно"
          .ErrorMessage = "Необходимо выбрать значение из списка "
          .IgnoreBlank = True: .InCellDropdown = True
          .Parent.Locked = False
        End With
      End If
      ProtectSheet Sh ' rev.330
    End If
    
    PartNumRow = Target.Row ' rev.380
    If (Target.Column = 6 Or Target.Column = 7) And PartNumRow > 1 Then
      If Sh.Range(Clipboard_RangeAddress).Column > 5 Then
        App.ScreenUpdating = False ' ВЫКЛ Обновление экрана
        If Not IsEmpty(Sh.Cells(Target.Row, 5)) Then _
          CostUpdate Sh.Cells(Target.Row, 5) ' Костыль And SuppNumRow > 0
      End If ': cnfRenew = Sh.Name ' rev.410
    Else: If CostChanged Then CostUpdate ' rev.410 - CostChanged ' Необходимость обновить статистику rev.340
    End If: cnfRenew = Sh.Name ' rev.410
  End If: AutoOpenControls ' rev.330
  'Application.ScreenUpdating = True ' ВКЛ Обновление экрана
  Clipboard_RangeAddress = Selection.Address ' Для буфера
End Sub
