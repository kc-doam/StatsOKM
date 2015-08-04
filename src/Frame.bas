Attribute VB_Name = "Frame"
Option Explicit
Option Base 1
'Option Private Module ' rev.340
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890

Private Counter As Integer ' Счётчик

Property Get GetUserName() As String
  GetUserName = Environ("UserName")
End Property

Property Let Quit(ByVal xlBlock As Boolean) ' вместо End rev.340
  With Application
    If xlBlock Then
      With UnprotectSheet(App_Wb.Sheets(GetSheetList(Set_spName)))
        .Cells.Locked = True
      End With: ProtectSheet App_Wb.Sheets(Sh_List(Set_spName))
      .CellDragAndDrop = True: .MoveAfterReturnDirection = xlDown: End
    Else
      .CellDragAndDrop = False: .MoveAfterReturnDirection = xlToRight
      ActiveWindow.Caption = ActiveWorkbook.Name & " (rev." & revFile & ")" _
        & IIf(ActiveWorkbook.ReadOnly, "  [Только для чтения]", "") ' rev.360
    End If
  End With
End Property

' Загрузка данных с настройками
Public Sub SettingsStatistics(ByRef Settings As Collection) ' rev.300
Dim iND As Object, Bank As String, SubBank As String
Const Let_accPath = "X:\Avtor_M\#Finansist\YCHET" ' Директория «YCHET» rev.330
  ' ВАЖНО! Обновление списка с Индексами листов
  If GetSheetList(Set_cnfName) < 1 Then ErrCollection 1001, 1, 16 ' EPN = 1
'  Worksheets(Sh_List(Set_cnfName)).Visible = xlSheetVeryHidden ' СКРЫТЬ rev.330
  RemoveCollection Settings: Settings.Add "#1/1/2009#", "date0" ' для SQL
  For Each iND In App_Wb.Sheets(Sh_List(Set_cnfName)).NameS ' Из листа «Настройки»
    With iND
      Bank = Left(.Name, InStr(.Name, "_")): SubBank = Mid(.Name, Len(Bank) + 1)
      If Not .Value Like "*[#]*" Then
        If Bank Like "*!" & Set_cnfName And .RefersToRange.Count = 1 Then _
          Settings.Add CStr(.RefersToRange.Value), SubBank
      Else: ErrCollection 57, 1, 16, "'" & .Name & "'": End If ' EPN = 1
    End With
  Next iND: Settings.Add IIf(Len(Dir(Let_accPath, vbDirectory)) > 0, _
    Let_accPath, ActiveWorkbook.Path), "SetPath" ' rev.330
End Sub

' Обновление списка Индексов листов
Public Function GetSheetList(ByVal FindCodeNameSheet As String) As Byte ' rev.300
Dim App_Sh As Worksheet: RemoveCollection Sh_List
  On Error Resume Next
    For Each App_Sh In App_Wb.Sheets
      Sh_List.Add App_Sh.Index, App_Sh.CodeName ' Добавляем индекс в список
      If App_Sh.CodeName = FindCodeNameSheet _
      Or App_Sh.Name = FindCodeNameSheet Then GetSheetList = App_Sh.Index
    Next App_Sh
    If Err Then ErrCollection Err.Number, 2, 16: Quit = True ' EPN = 2
End Function

Public Sub ProtectSheet(ByRef Sh As Worksheet) ' Защитить лист
  On Error Resume Next
    Sh.EnableOutlining = True ' ЗАГРУЗКА: Группировка на защищённом листе
    Sh.Protect Password:=Settings("CostPass"), UserInterfaceOnly:=True, _
      Contents:=True, AllowFiltering:=True, AllowDeletingRows:=True, _
      AllowFormattingColumns:=True, DrawingObjects:=False
    If Err Then ErrCollection Err.Number, 2, 16, Sh.Name ' EPN = 2
End Sub

' Снять защиту с листа
Public Function UnprotectSheet(ByRef Sh As Worksheet) As Worksheet
  On Error Resume Next
    If Sh.ProtectScenarios Then Sh.Unprotect Settings("CostPass")
    Set UnprotectSheet = Sh
    If Err Then ErrCollection Err.Number, 2, 16, Sh.Name ' EPN = 2
End Function

Public Sub SortSupplier(ByRef Sh As Worksheet, _
ByVal FirstKey As Byte, Optional ByVal SecondKey As Byte)
Dim LastRow As Long: LastRow = Sh.UsedRange.Rows.Count + 1 ' Последняя строка
  If Not Sh.AutoFilterMode Then Sh.Cells(1, FirstKey).AutoFilter
  With Sh.AutoFilter.Sort
    .SortFields.Clear: .Header = xlYes
    .SortFields.Add Key:=Sh.Cells(2, FirstKey).Resize(LastRow, 1)
    If SecondKey > 0 Then _
      .SortFields.Add Key:=Sh.Cells(2, SecondKey).Resize(LastRow, 1)
    .Orientation = xlTopToBottom: .Apply
  End With
End Sub

' Удаление коллекции
Private Sub RemoveCollection(ByRef CollectionName As Collection) ' rev.300
  For Counter = 1 To CollectionName.Count: CollectionName.Remove 1: Next Counter
End Sub

' Заполнение одномерного массива из двумерного (для цен)
Public Function MultidimArr(ByVal Cost As Variant, ByVal Row As Long, _
Optional ByRef FirstItem As Byte = 0) As Variant
Dim Arr As Variant
  If FirstItem > 0 Then ' для Цен
    ReDim Arr(LBound(Cost, 1) + FirstItem To UBound(Cost, 1)) As Currency
    If Not IsArray(Cost) Then MultidimArr = Arr: Exit Function
    For Counter = LBound(Arr) To UBound(Arr)
      If Not IsNull(Cost(Counter, Row)) Then Arr(Counter) = Cost(Counter, Row)
    Next Counter: MultidimArr = Arr
  Else ' для текстовых массивов
    ReDim Arr(LBound(Cost, 2) To UBound(Cost, 2)) As String
    For Counter = LBound(Arr) To UBound(Arr)
      Arr(Counter) = Cost(Row, Counter)
    Next Counter: MultidimArr = Arr
  End If
End Function

' Указания для пользователя при возникновении ошибки
Public Sub ErrCollection(ByVal ErrNumber As Long, ByVal ErrPartNum As Byte, _
ByVal Icon As Byte, Optional ByVal Str As String)
Dim Ask As Byte, Msg As String, Title As String:
  Ask = 1: Title = "Ошибка чтения " ' По умолчанию
  Select Case ErrNumber * ErrPartNum ' Номер ошибки * EPN (ErrPartNum)
    ' EPN = 1
    Case -2147217843: Msg = "Неверный пароль базы данных. " _
      & "Восстановите резервную копию файла " & vbCrLf & Str
    Case 20: Ask = 0: Msg = "У поставщика '" & Str & "' изменились основные " _
      & "данные. " & vbCrLf & "Перед сохранением необходимо изменить поле " _
      & "'Дата актуальности'. " & vbCrLf: Title = "Ошибка ввода данных "
    Case 30: Ask = 2: Msg = "Внимание! Обновился файл с ценами. ": Title = _
      "Требуется обновление "
    ' В данной версии нет предупреждения «Дата поступления» с пустым поставщиком
    Case 40: If Str Like "*''*" Then _
      Ask = 5: Msg = "Не указан поставщик " & Mid(Str, 19) & ". ": Icon = 64 _
      Else: Ask = 4: Msg = "Не найдены цены " & Str & ". " ' rev.340
    Case 57: Msg = "В настройках " & Str & " обнаружена битая ссылка. "
    Case 59: Msg = "Файл '" & Str & "' не найден! " _
      & "Работа с данными невозможна!": Title = "Ошибка открытия файла "
    Case 457: Ask = 2: Msg = "Невозможно обновить коллекцию с ценами '" & Str _
      & "'. Работа с данными невозможна! "
    Case 1001: Ask = 3: Msg = "Лист 'Настройки' не найден! " _
      & "Работа с данными невозможна! "
    ' EPN = 2
    Case 10: Ask = 2: Msg = IIf(Len(Str), "Невозможно снять защиту с листа '" _
      & Str & "'. ", "Лист не защищён. ") & "Коллекция 'Settings' is Nothing! "
    Case 182, 184: Ask = 0: Msg = "Значение переменной 'App_Wb' is Nothing! " _
      & "Работа с данными невозможна! " & vbCrLf & IIf(ErrNumber = 92, "Необ" _
      & "ходимо сохранить файл '" & Windows(1).Caption & "' и открыть заново. " _
      & vbCr & vbCrLf & "При частом появлении ошибки о", "О") & "братитесь " _
      & "к специалисту по автоматизации. ": Title = "Внутренняя ошибка "
    Case 2008: Ask = 3: Msg = "На листе '" & Str & "' задан неизвестный пароль. "
    ' EPN = 3
    Case 21: Msg = "Ошибка в формуле условного форматирования на листе '" _
      & Str & "'. ": Title = "Ошибка ввода данных " ' rev.360
    Case 273: Msg = "Невозможно применить сортировку к пустому фильтру " _
      & "на листе '" & Str & "'. "
    Case 3012: Msg = "Невозможно применить автофильтр на листе '" & Str & "'. "
    Case 3018: Msg = "Невозможно создать условное форматирование. Ошибка " _
      & "в связанных диапазонах, либо лист '" & Str & "' защищён от записи. " _
      & vbCrLf: Title = "Ошибка ввода данных "
    ' not EPN
    Case Else: Msg = "Неизвестная ошибка #" & ErrNumber & " ": Icon = 16
  End Select: Select Case Ask
    Case 1: Msg = Msg & vbCrLf & "Обратитесь к специалисту по автоматизации. "
    Case 2: Msg = Msg & vbCrLf & "Необходимо сохранить файл '" _
      & Windows(1).Caption & "' " & "и открыть заново. "
    Case 3: Msg = Msg & vbCrLf & "Восстановите резервную копию " _
      & "файла '" & Windows(1).Caption & "'. ": Title = "Критическая ошибка "
    Case 4: Msg = Msg & vbCrLf & "Проверьте 'Категорию цен' у поставщика, " _
      & "затем проставьте 'Дату поступления в ОКМ'. "
    Case 5: Msg = Msg & vbCrLf & "Выберите поставщика " _
      & "или удалите 'Дату поступления в ОКМ'. "
  End Select: MsgBox Msg, Icon, Title & IIf(ErrNumber > 0, ErrPartNum & "x", _
    "ADODB ") & ErrNumber: If Ask = 3 Then Quit = True
End Sub
