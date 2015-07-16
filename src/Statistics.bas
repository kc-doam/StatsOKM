Attribute VB_Name = "Statistics"
Option Explicit
Option Base 1
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890
Public Const Let_SuppList = "Список_Поставщиков", revFile As Integer = 330
' Внутреннее имя листа «Архив», Внутреннее имя листа «Поставщики»
Public Const Set_arName = "ARCH_", Set_spName = "SUPP_"
Public Const Let_ContrFormList = "станд.,опер."
Private Const Let_OrgBodyList = "Ф/Л,Ю/Л" ' Не менять!

' Рабочая книга, Рабочий лист, Имя рабочего листа/параметра
Public App_Wb As Workbook, cnfRenew As String
' Массив с изменениями о поставщике, Номер выделенной строки о поставщике
Public SuppDiff As Variant, SuppNumRow As Long, PartNumRow As Long

' Коллекция с настройками книги, Коллекция с именами листов
Public Settings As New Collection, Sh_List As New Collection
' Коллекция с ценами ОКМ для Ф/Л и Ю/Л, Счётчик
Private Cost As New Collection, Counter As Integer
' Коллекции: ключи коллекции BankID, рабочие листы и колонки
'Public BankID As New Collection
' Коллекция: реквизиты и контакты поставщиков, цены
'Public BankSUPP As New Collection

Private Sub Auto_Open() ' Автомакрос
Dim cstPath As String, mdwPath As String
Dim Conn As Object, Rec As Object, Src As String
Const Let_mdwPath = "\Application Data\Microsoft\Access\System.mdw"
Const Let_cstPath = "\Архив\Cost.accdb" ' Цены
  SettingsStatistics Settings ' Загрузка настроек книги в коллекцию
  PartNumRow = ActiveCell.Row ' Номер строки партии материалов
  
  ' Проверка существования директории с настройками rev.330
  Set Conn = CreateObject("Scripting.FileSystemObject") ' fso
  cstPath = Settings("SetPath") & Let_cstPath: If Not Conn _
    .FileExists(cstPath) Then ErrCollection 59, 1, 16, cstPath: End ' EPN = 1
  Debug.Print Settings("SetPath"): Set Conn = Nothing
  ' Создание системной таблицы, если она не существует
  Set Conn = CreateObject("Scripting.FileSystemObject") ' fso
  mdwPath = Environ("UserProfile") & Let_mdwPath ' Полный путь к файлу
  Do While Not Conn.FileExists(mdwPath) ' Выполнять ПОКА нет файла
    mdwPath = Left(mdwPath, InStrRev(mdwPath, "\") - 1)
    Do Until Conn.FolderExists(mdwPath) ' Выполнять ДО того, как появится папка
      Do While Not Conn.FolderExists(mdwPath) ' Выполнять ПОКА нет папки
        Src = Right(mdwPath, Len(mdwPath) - InStrRev(mdwPath, "\"))
        mdwPath = Left(mdwPath, InStrRev(mdwPath, "\") - 1)
      Loop: Conn.GetFolder(mdwPath).SubFolders.Add Src ' Создать папку
      mdwPath = Environ("UserProfile") & Let_mdwPath ' Полный путь к файлу
      mdwPath = Left(mdwPath, InStrRev(mdwPath, "\") - 1)
    Loop
    mdwPath = Environ("UserProfile") & Let_mdwPath ' Полный путь к файлу
    ByteArrayToSystemMdw mdwPath ' Создаём системную таблицу
  Loop: Set Conn = Nothing
  
  On Error Resume Next
    Set Conn = CreateObject("ADODB.Connection") ' Открываем Connection
    Conn.ConnectionTimeout = 5
    Conn.Mode = 1 ' 1 = adModeRead, 2 = adModeWrite, 3 = adModeReadWrite
    Src = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cstPath & ";"
    'Src = Src & "Jet OLEDB:Engine Type=6;" ' Тип подключения
    Src = Src & "Jet OLEDB:Encrypt Database=True;"
    Src = Src & "Jet OLEDB:Database Password=" & Settings("CostPass") & ";"
    Src = Src & "Jet OLEDB:System database=" & mdwPath ' Системная таблица
    Conn.Open ConnectionString:=Src ', UserId:="admin", Password:=""
    If Err Then ErrCollection Err.Number, 1, 16, cstPath: End ' EPN = 1
  On Error GoTo 0
  
  Set Rec = CreateObject("ADODB.Recordset") ' Создаём RecordSet
  With Rec ' Загружаем Цены в коллекцию
    For Each SuppDiff In Split(Let_OrgBodyList, ",")
      cnfRenew = SuppDiff: Src = "SELECT Name FROM [MSysObjects] " _
        & "WHERE Flags = 0 AND Type = 1 AND Name LIKE '" & cnfRenew & "%' "
      .Open Source:=Src, ActiveConnection:=Conn
      Src = Empty: SuppDiff = .GetRows(Rows:=-1): .Close
      For Counter = LBound(SuppDiff, 2) To UBound(SuppDiff, 2)
      'Src=Src & "SELECT * FROM [" & SuppDiff(LBound(SuppDiff, 1), Counter) & "]"
        Src = Src & "SELECT MID('" & SuppDiff(LBound(SuppDiff, 1), Counter) _
          & "', 5) AS 'Имя таблицы', Актуально, [Группа 0], [Группа 1], [Груп" _
          & "па 2], [Группа А], [НУМ 0], [НУМ 1], [НУМ 2], [НАШ 1], [НАШ 2], " _
          & "Вопросы" & IIf(cnfRenew = "Ф/Л", ", [Оф письма], Бухонлайн, Коде" _
          & "кс", "") & " FROM [" & SuppDiff(LBound(SuppDiff, 1), Counter) & "]"
        ' !Последний UNION ALL - НЕ ВКЛЮЧАТЬ!
        If Counter < UBound(SuppDiff, 2) Then Src = Src & " UNION ALL "
      Next Counter
      On Error Resume Next
        ' При открытии пустого объекта Recordset свойства BOF и EOF содержат True
        .Open Source:=Src, ActiveConnection:=Conn
        Cost.Add .GetRows(Rows:=-1), cnfRenew: .Close
      Debug.Print cnfRenew & " - Err #" & Err.Number; "; Fields "; .Fields.Count
        If Err.Number = 457 Then
          If Err Then ErrCollection Err.Number, 1, 16, cnfRenew: End ' EPN = 1
        ElseIf Err.Number = 3021 Then
          Src = "SELECT 'стандарт' AS '0', " & Settings("date0") & " AS '1'"
          For Counter = 2 To .Fields.Count - 1
            Src = Src & ", " & IIf(Counter < 5, "-1", "NULL") _
              & " AS '" & Counter & "'"
          Next Counter: .Open Source:=Src, ActiveConnection:=Conn
          Cost.Add .GetRows(Rows:=-1), cnfRenew: .Close
        End If: cnfRenew = ""
      On Error GoTo 0
    Next SuppDiff
  End With: Set Rec = Nothing: Set Conn = Nothing
  
  Set Conn = CreateObject("Scripting.FileSystemObject") ' fso
    Debug.Print Mid(Log(Conn.GetFile(cstPath).DateLastModified) - 10, 3, 8)
    'MsgBox Mid(Log(Conn.GetFile(cstPath).DateLastModified) - 10, 3, 8)
  Set Conn = Nothing
  
  App_Wb.Sheets(Sh_List("SF_")).Activate ' ВАЖНО! Уйти с листа «Поставщики»
  For Each Rec In App_Wb.Names ' Диапазон «Список_Поставщиков»
    If Rec.Name = Let_SuppList Then Rec.Delete
  Next Rec
  With ActiveWorkbook.Names.Add(Name:=Let_SuppList, RefersTo:="=$E$1")
    .Comment = "Список листа Поставщики": .RefersTo = "=OFFSET('" _
      & App_Wb.Sheets(Sh_List(Set_spName)).Name & "'!$J$1,1,,COUNTA('" _
      & App_Wb.Sheets(Sh_List(Set_spName)).Name & "'!$J:$J)-1,)"
  End With
  If Not App_Wb.ReadOnly Then
    For Each Conn In App_Wb.Sheets
'      UnprotectSheet Conn: ProtectSheet Conn ' Временные параметры rev.330
      ProtectSheet Conn ' Временные параметры rev.330
    Next Conn: cnfRenew = "": App_Wb.Save ' Если доступен, сохранить
  End If: cnfRenew = ActiveSheet.Name
End Sub

' Установка рабочей конфигурации листов
Public Sub SpecificationSheets(ByVal SheetIndex As Byte)
Dim App_Sh As Worksheet, LastRow As Long
  'Stop ' #3 Копирование заголовка, Установка условного форматирования
  Application.ScreenUpdating = False ' ВЫКЛ Обновление экрана
  For Each App_Sh In App_Wb.Sheets
    With UnprotectSheet(App_Wb.Sheets(App_Sh.Index))
      .Activate ': UnprotectSheet App_Sh
      With ActiveWindow ' CTRL+HOME
        .ScrollRow = 1: .ScrollColumn = 1: .FreezePanes = False
      End With
      
      LastRow = App_Sh.UsedRange.Rows.Count ' Количество строк rev.330
      ' Очистка форматов, Очистка условного форматирования
      .Cells.ClearFormats: .Cells.FormatConditions.Delete
      ' Очистка группировки, Очистка проверки данных
      .Cells.ClearOutline: .Cells.Validation.Delete
      
      Select Case .CodeName
        Case Set_spName, Set_arName ' Лист «Поставщики», «Архив»
          If .CodeName = Set_arName Then App_Wb.Sheets(Sh_List(Set_spName)) _
            .Range("A1:O1").Copy Destination:=.Range("A1")
          On Error Resume Next
            'SendKeys "^{HOME}", False ' rev.250 Фокус должен быть на MS Excel
            '.Cells(1, 1).AutoFilter
            
            If .AutoFilterMode Then .ShowAllData Else .Cells(1, 1).AutoFilter ' Автофильтр
            
            If Err Then ErrCollection Err.Number, 3, 48, .Name ' EPN = 3
          On Error GoTo 0
          .Columns("C:I").Columns.Group: .Columns("F:G").Columns.Group
          If .CodeName = Set_spName Then
            .Columns("A:AB").Locked = False: .Rows("1:1").Locked = True
            .Columns("N:X").Columns.Group
            ' Закрепление области
            .Range("K2").Select: ActiveWindow.FreezePanes = True
          End If
          ' Форматирование колонок
          .Columns("D:D").NumberFormat = "@"
          .Columns("K:K").NumberFormat = "m/d/yyyy"
          If .CodeName = Set_spName Then
            .Columns("Q:Q").NumberFormat = "m/d/yyyy"
            .Columns("R:R").NumberFormat = "[$-419]"".+. (""0000"") "";@"
            .Columns("V:V").NumberFormat = "[$-419]000-000-000-00;@"
            .Columns("W:W").NumberFormat = _
              "[<=9999999999]0000000000;000000000000;@"
            .Columns("X:X").NumberFormat = "[$-419]000000000;@"
            .Columns("AA:AA").NumberFormat = "@"
          End If
          ' Сортировка
          SortSupplier App_Sh, 10, 11
          ' Условное форматирование
          If Val(Application.Version) >= 12 Then
            With .Range("A2:A9,D2:E9,L2:M9").FormatConditions _
              .Add(Type:=xlBlanksCondition)
              .Interior.ColorIndex = 3: .StopIfTrue = True
            End With
            With .Range("C2:C9").FormatConditions _
              .Add(Type:=xlBlanksCondition)
              .Interior.ColorIndex = 44: .StopIfTrue = True
            End With
            ' Поставщик (кратко)
            With .Range("J2:K9").FormatConditions _
              .Add(Type:=xlExpression, Formula1:="=И(НЕ(ЕПУСТО($J2));" _
                & "ИЛИ(ЕПУСТО($A2);ЕПУСТО($L2);ЕПУСТО($M2)))")
              .Font.ColorIndex = 2: .Interior.ColorIndex = 9
              .StopIfTrue = True: .SetFirstPriority
            End With
            If .CodeName = Set_spName Then
              ' ИНН
              With .Range("W2:W9").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=ИЛИ(ЕСЛИ(НЕ(ЕЧИСЛО($W2));" _
                  & "1;ЦЕЛОЕ($W2)<>$W2);ИЛИ(И($A2=""Ф/Л"";ДЛСТР($W2)<11));" _
                  & "И($A2=""Ю/Л"";ДЛСТР($W2)>10))")
                .Interior.ColorIndex = 44: .StopIfTrue = True
              End With
              With .Range("J2:J9").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=И(НЕ(ЕПУСТО(СМЕЩ($J2;-1;" _
                  & "0)));ЕПУСТО($J2))")
                .Interior.ColorIndex = 36: .StopIfTrue = True
              End With
              .Range("O2:AB9").FormatConditions.Add Type:=xlNoBlanksCondition
              With .Range("O2:O9,U2:U9,Z2:Z9").FormatConditions _
                .Add(Type:=xlBlanksCondition)
                .Interior.ColorIndex = 36: .StopIfTrue = True
              End With
              With .Range("Q2:Q9,T2:T9,V2:V9").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=$A2=""Ф/Л""")
                .Interior.ColorIndex = 36: .StopIfTrue = True
              End With
              With .Range("N2:N9,P2:P9,X2:X9").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=$A2=""Ю/Л""")
                .Interior.ColorIndex = 36: .StopIfTrue = True
              End With
            End If
            If .CodeName = Set_arName Then
              With .Range("A2:O9").FormatConditions _
                .Add(Type:=xlExpression, _
                  Formula1:="=ЕПУСТО($J2)+МАКС(--($J2:$J9=$J2)*$K2:$K9)=$K2")
                .Interior.ColorIndex = 43: .StopIfTrue = True
                '.SetFirstPriority
              End With
              With .Range("A2:O9").FormatConditions _
                .Add(Type:=xlExpression, _
                  Formula1:="=И($K2>СЕГОДНЯ()-90;НЕ(ЕПУСТО($J2)))")
                .Interior.ColorIndex = 27: .StopIfTrue = True
                '.SetFirstPriority
              End With
            End If
          End If
          ' Проверка ввода данных
          If .CodeName = Set_spName Then
            With .Range("A2:A9").Validation
              .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Formula1:=Let_OrgBodyList
              .ErrorTitle = "Вид лица"
              .ErrorMessage = "Необходимо выбрать значение из списка "
              .ShowError = True: .IgnoreBlank = True
            End With
            With .Range("D2:D9").Validation
              .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, _
                Formula1:="=OR(AND($D2>=""001"",$D2<=""999"",LEN($D2)<4)," _
                & "$D2=""станд."",$D2=""ДПР"")"
              .ErrorTitle = "Источник"
              .ErrorMessage = "Необходимо ввести 3-х значный номер РИЦа, " _
                & "либо указать источник ""ДПР"" или ""станд."""
              .ShowError = True: .IgnoreBlank = True
            End With
            With .Range("E2:E9").Validation
              .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Formula1:="Коммерч.,Некоммерч.,Ведомство,РИЦ," _
                & "Собств. источники КЦ"
              .ErrorTitle = "Тип организации"
              .ErrorMessage = "Необходимо выбрать значение из списка "
              .ShowError = True: .IgnoreBlank = True
            End With
            With .Range("K2:K9").Validation
              .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
                Operator:=xlGreaterEqual, Formula1:=Replace( _
                Settings("date0"), "#", "")
              .ErrorTitle = "Дата актуальности"
              .ErrorMessage = "Необходимо ввести дату не раньше " & .Formula1
              .ShowError = True: .IgnoreBlank = True
            End With
            With .Range("M2:M9").Validation
              .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Formula1:="НДС,УСН"
              .ErrorTitle = "НДС / УСН"
              .ErrorMessage = "Необходимо выбрать значение из списка "
            End With
            With .Range("Q2:Q9").Validation
              .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
                Operator:=xlGreaterEqual, Formula1:=DateAdd( _
                "m", -840, Replace(Settings("date0"), "#", ""))
              .ErrorTitle = "Дата рождения"
              .ErrorMessage = "Необходимо ввести дату не раньше " & .Formula1
              .ShowError = True: .IgnoreBlank = True
            End With
            With .Range("R2:R9").Validation
              .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlGreaterEqual, Formula1:=DatePart( _
                "yyyy", Replace(Settings("date0"), "#", ""))
              .ErrorTitle = "Заявление о проф. вычете"
              .ErrorMessage = "Необходимо ввести год не меньше " & .Formula1
              .ShowError = True: .IgnoreBlank = True
            End With
            With .Range("V2:V9").Validation
              .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, Formula1:="100000000", _
                Formula2:="99999999999"
              .ErrorTitle = "СНИЛС"
              .ErrorMessage = "Страховой номер в пенсионном фонде должен " _
                & "содержать от 9 до 11 цифр "
              .ShowError = True: .IgnoreBlank = True
            End With
            With .Range("W2:W9").Validation ' ИНН
              .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, _
                Formula1:="=AND(OR(AND($A2=""Ф/Л"",LEN($W2)<13)," _
                & "AND($A2=""Ю/Л"",LEN($W2)<11)),LEN($W2)>8)"
              .ErrorTitle = "ИНН"
              .ErrorMessage = "Идентификационный номер налогоплательщика " _
                & "должен содержать: " & vbCrLf & vbTab & "для Ф/Л  от 11 " _
                & "до 12 цифр" & vbCrLf & vbTab & "для Ю/Л  от 9 до 10 цифр"
              .ShowError = True: .IgnoreBlank = True
            End With
            With .Range("X2:X9").Validation
              .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, Formula1:="10000000", Formula2:="999999999"
              .ErrorTitle = "КПП"
              .ErrorMessage = "Код причины постановки должен " _
                & "содержать от 8 до 9 цифр "
              .ShowError = True: .IgnoreBlank = True
            End With
            ' Список «Категория цены» для Поставщика
            'ListCost Sh_List(Set_spName), 1 ' ОТМЕНИТЬ ПОВТОРНОЕ СОЗДАНИЕ СПИСКА
          End If
        Case "SF_", "SB_"
          .Tab.ColorIndex = 24
          'On Error Resume Next
            'SendKeys "^{HOME}", False ' rev.250 Фокус должен быть на MS Excel
            '.Cells(1, 1).AutoFilter
            
            If .AutoFilterMode Then .ShowAllData Else .Cells(1, 1).AutoFilter ' Автофильтр
            
            If Err Then ErrCollection Err.Number, 3, 48, .Name ' EPN = 3
          'On Error GoTo 0
          .Columns("E:I").Locked = False: .Columns("K:R").Locked = False
          .Columns("T:AA").Locked = False ' rev.330
          .Columns("AC:AE").Locked = False: .Columns("AG:AO").Locked = False
          .Columns("AQ").Locked = False: .Columns("AS").Locked = False
          .Columns("AV:AW").Locked = False: .Rows("1:1").Locked = True
          .Columns("B:D").Columns.Group: .Columns("I:R").Columns.Group
          .Columns("Y:AA").Columns.Group: .Columns("AC:AE").Columns.Group
          .Columns("AG:AQ").Columns.Group: .Columns("AL:AO").Columns.Group
          ' Закрепление области
          .Range("G2").Select: ActiveWindow.FreezePanes = True
          With Selection ' Выделяем ячейку в последней строке
            If .Row < LastRow Then .Offset(, -2).End(xlDown).Select ' rev.330
          End With
          ' Форматирование колонок
          .Columns("A:D").NumberFormat = "General"
          .Range("F:O,Q:Q,T:X").NumberFormat = "m/d/yyyy"
          'With
          '  .NumberFormat = "#,##0"
          '  .HorizontalAlignment = xlRight
          '  .IndentLevel = 1
          'End With
      End Select: ProtectSheet App_Sh
    End With
  Next App_Sh
  'On Error Resume Next
  '  App_Wb.Sheets(SheetIndex).Select
  '  If Err Then App_Wb.Sheets(Sh_List("SF_")).Select
  App_Wb.Sheets(Sh_List("SF_")).Select
  'On Error GoTo 0
  Application.ScreenUpdating = True ' ВКЛ Обновление экрана
  'SendKeys "{NUMLOCK}", True ' Костыль v2.5 Фокус должен быть на MS Excel
End Sub

' Запись в «Архив» данных о поставщике из массива SuppDiff
Public Sub RecordCells(ByVal NewSupplier As Boolean)
Dim i As Integer: Counter = 0: i = 2
  ''Stop
  If NewSupplier And IsArray(SuppDiff) Then ' => SuppNumRow = 0
    On Error Resume Next
      If Len(SuppDiff(10)) > 0 And Len(SuppDiff(11)) > 0 Then
        With UnprotectSheet(App_Wb.Sheets(Sh_List(Set_arName)))
        'With App_Wb.Sheets(Sh_List(Set_arName))
          Do Until IsEmpty(.Cells(i, 10)) ' Счётчик строк ' Выполнять ДО
            ' Поставщик без «Даты актуальности» не добавляется
            If .Cells(i, 10) = SuppDiff(10) _
            And .Cells(i, 11) = CDate(SuppDiff(11)) Then Counter = i
            i = i + 1
            Debug.Print "RecordCells Err #" & Err.Number: If Err Then Exit Sub
          Loop: If Counter <> 0 Then i = Counter
          
          .Cells(i, 1).Resize(, UBound(SuppDiff)) = SuppDiff
          If IsDate(SuppDiff(11)) Then .Cells(i, 11) = CDate(SuppDiff(11))
          
          SortSupplier App_Wb.Sheets(Sh_List(Set_arName)), 10, 11
          
          If Err Then ErrCollection Err.Number, 3, 16, .Name ' EPN = 3
        End With: SuppNumRow = 0: SuppDiff = Empty ' Очищаем массив SuppDiff
        ProtectSheet App_Wb.Sheets(Sh_List(Set_arName))
      End If
    On Error GoTo 0 ' ВАЖНО! Отключаем сообщения об ошибках
  End If
End Sub

' Изменились ли данные о поставщике на листе «Поставщики»
Public Function CheckSupplier() As Boolean
  On Error Resume Next
    ' ВАЖНО! Обновление списка с Индексами листов
    With App_Wb.Sheets(GetSheetList(Set_spName))
      If Not .Cells(SuppNumRow, 1).Value & .Cells(SuppNumRow, 4).Value _
        & .Cells(SuppNumRow, 5).Value & .Cells(SuppNumRow, 6).Value _
        & .Cells(SuppNumRow, 8).Value & .Cells(SuppNumRow, 10).Value _
        & .Cells(SuppNumRow, 11).Value & .Cells(SuppNumRow, 12).Value _
        & .Cells(SuppNumRow, 13).Value & .Cells(SuppNumRow, 15).Value _
        = SuppDiff(1) & SuppDiff(4) & SuppDiff(5) & SuppDiff(6) & SuppDiff(8) _
        & SuppDiff(10) & SuppDiff(11) & SuppDiff(12) & SuppDiff(13) _
        & SuppDiff(15) And Len(.Cells(SuppNumRow, 11).Value) > 0 Then
        
        If Err Then ErrCollection 10, 1, 16: Exit Function ' EPN = 1
        
        If .Cells(SuppNumRow, 10).Value & .Cells(SuppNumRow, 11).Value _
          = SuppDiff(10) & SuppDiff(11) Then
          .Activate
          If MsgBox("У поставщика '" & SuppDiff(10) & "' изменились основные " _
            & "данные. Необходимо изменить 'Дату актуальности'. " & vbCrLf _
            & "Изменить 'Дату актуальности' " & SuppDiff(11) & " на " _
            & "текущую дату? ", 260 + 48, "Данные о поставщике") = vbYes Then
            .Cells(SuppNumRow, 11) = Date
          Else
            .Cells(SuppNumRow, 11).Select: Exit Function
          End If
        End If
        CheckSupplier = True ' Подтвердить изменение данных
        ' Создаём массив с изменениями о поставщике rev.310
        SuppDiff = MultidimArr(.Cells(SuppNumRow, 1) _
          .Resize(, UBound(SuppDiff)).Value, 1)
      End If
    End With
End Function

' Проверка перед действиями App_WorkbookBeforeClose и App_WorkbookBeforeSave
Public Function ChangedBeforeSave(ByRef Sh As Worksheet) As Boolean
  If Sh.CodeName = Set_spName And IsArray(SuppDiff) And SuppNumRow > 1 Then
    ''Stop
    If Not CheckSupplier Then ' Если изменены данные о поставщике
      'If Len(SuppDiff(11)) > 0 And Not Sh.Cells(SuppNumRow, 1).Value _
        & Sh.Cells(SuppNumRow, 4).Value & Sh.Cells(SuppNumRow, 5).Value _
        & Sh.Cells(SuppNumRow, 6).Value & Sh.Cells(SuppNumRow, 8).Value _
        & Sh.Cells(SuppNumRow, 10).Value & Sh.Cells(SuppNumRow, 12).Value _
        & Sh.Cells(SuppNumRow, 13).Value & Sh.Cells(SuppNumRow, 15).Value _
        = SuppDiff(1) & SuppDiff(4) & SuppDiff(5) & SuppDiff(6) & SuppDiff(8) _
        & SuppDiff(10) & SuppDiff(12) & SuppDiff(13) & SuppDiff(15) Then
      If Not Sh.Cells(SuppNumRow, 1).Value & Sh.Cells(SuppNumRow, 4).Value _
        & Sh.Cells(SuppNumRow, 5).Value & Sh.Cells(SuppNumRow, 6).Value _
        & Sh.Cells(SuppNumRow, 8).Value & Sh.Cells(SuppNumRow, 10).Value _
        & Sh.Cells(SuppNumRow, 11).Value & Sh.Cells(SuppNumRow, 12).Value _
        & Sh.Cells(SuppNumRow, 13).Value & Sh.Cells(SuppNumRow, 15).Value _
        = SuppDiff(1) & SuppDiff(4) & SuppDiff(5) & SuppDiff(6) & SuppDiff(8) _
        & SuppDiff(10) & SuppDiff(11) & SuppDiff(12) & SuppDiff(13) _
        & SuppDiff(15) And Len(Sh.Cells(SuppNumRow, 11).Value) > 0 Then
        
        ErrCollection 20, 1, 16, Sh.Cells(SuppNumRow, 10) ' EPN = 1
        ChangedBeforeSave = True
      End If
    End If
  End If
End Function

' Список с «Категориями цены»
Public Sub ListCost(ByRef Sh As Worksheet, ByVal TargetRow As Long)
Dim Src As String, OrgBody As String ''' ???
  OrgBody = Sh.Cells(TargetRow, 1) ' Список «Категория цены» для Поставщика
  ''Stop ' Добавить OrgBody
  Sh.Cells(TargetRow, 12).Validation.Delete ' Очистка проверки данных
  
  If Len(OrgBody) > 2 And Let_OrgBodyList Like "*" & OrgBody & "*" Then
    For Counter = LBound(Cost(OrgBody), 2) To UBound(Cost(OrgBody), 2)
      If Cost(OrgBody)(LBound(Cost(OrgBody), 1), Counter) = "РИЦ" Then _
        Counter = Counter + 1 ' Пропускаем архивные цены (РИЦ до 2012 года)
      If Counter > LBound(Cost(OrgBody), 2) Then ' Пропускаем 1-ю запись
        ' Если таблица «Категория цены» = «Поставщик» <> предыдущее знач.таблицы
        If Cost(OrgBody)(LBound(Cost(OrgBody), 1), Counter) = Sh.Cells( _
          TargetRow, 10) And Cost(OrgBody)(LBound(Cost(OrgBody), 1), Counter) _
          <> Cost(OrgBody)(LBound(Cost(OrgBody), 1), Counter - 1) Then _
          Src = Cost(OrgBody)(LBound(Cost(OrgBody), 1), Counter) & "," & Src
      Else
        Src = "стандарт,ДПР" ' Исключаем как архивные «Категория цены» = «РИЦ»
      End If
    Next Counter
    With Sh.Cells(TargetRow, 12).Validation
      .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=Src
      .ErrorTitle = "Категория цен"
      .ErrorMessage = "Необходимо выбрать значение из списка "
      .ShowError = True: .IgnoreBlank = True
    End With
  End If
End Sub

Public Function GetCosts(ByRef Sh As Worksheet, ByVal PartRow As Long, _
ByVal OrgBody As String) As String ''' ???
Dim SuppCost As Variant
  If IsEmpty(Cost) Then GoTo DateExit
  With App_Wb.Sheets(Sh_List(Set_arName))
    For Counter = LBound(Cost(OrgBody), 2) To UBound(Cost(OrgBody), 2)
    
    Debug.Print Cost(OrgBody)(0, Counter); "="; .Cells(SuppNumRow, 12)
    Debug.Print Cost(OrgBody)(1, Counter); "<="; Sh.Cells(PartRow, 6)
    ''Stop
      
      If Cost(OrgBody)(0, Counter) = .Cells(SuppNumRow, 12) _
      And Cost(OrgBody)(1, Counter) <= Sh.Cells(PartRow, 6) Then
        ' ВАЖНО! Если следующее поле цены «Актуально» > «Даты актуальности»
        If Cost(OrgBody)(0, Counter) <> Cost(OrgBody)(0, Counter + 1) _
        Or (Cost(OrgBody)(0, Counter) = Cost(OrgBody)(0, Counter + 1) _
        And Cost(OrgBody)(1, Counter + 1) > Sh.Cells(PartRow, 6)) Then _
          SuppCost = MultidimArr(Cost(OrgBody), Counter, 2): Exit For
      End If
    Next Counter: If Counter > UBound(Cost(OrgBody), 2) Then GoTo DateExit
  End With
  Select Case Sh.CodeName
    Case "SF_" ' КФ
      If Len(OrgBody) > 2 And Let_OrgBodyList Like "*" & OrgBody & "*" Then
        For Counter = LBound(SuppCost) To UBound(SuppCost)
          PartRow = SuppCost(Counter)
          Select Case Counter
            Case 2: GetCosts = "=RC[-11]*" & PartRow ' Группа 0
            Case 3: GetCosts = GetCosts & "+RC[-10]*" & PartRow ' Группа 1
            Case 4: GetCosts = GetCosts & "+RC[-9]*" & PartRow ' Группа 2
            Case 6: If PartRow > 0 Then GetCosts = Replace(GetCosts, _
              "RC[-11]", "(RC[-11]-RC[-6])") & "+RC[-6]*" & PartRow ' НУМ 0
            Case 7: If PartRow > 0 Then GetCosts = Replace(GetCosts, _
              "RC[-10]", "(RC[-10]-RC[-5])") & "+RC[-5]*" & PartRow ' НУМ 1
            Case 8: If PartRow > 0 Then GetCosts = Replace(GetCosts, _
              "RC[-9]", "(RC[-9]-RC[-4])") & "+RC[-4]*" & PartRow ' НУМ 2
            Case 9: GetCosts = GetCosts & "+RC[-8]*" & PartRow ' НАШ 1
            Case 10: GetCosts = GetCosts & "+RC[-7]*" & PartRow ' НАШ 2
            'Case 12: GetCosts = GetCosts & "+RC[-3]*" & PartRow ' Оф письма Ф/Л
            'Case 14: GetCosts = GetCosts & "+RC[-2]*" & PartRow ' Бухонлайн Ф/Л
            Case 13: If PartRow > 0 Then GetCosts = GetCosts _
              & "+RC[-1]*" & PartRow ' Кодекс Ф/Л
          End Select
        Next Counter
      End If
    'Case "" ' Актуализация материалов [Группа А] = 5
    'Case "" ' Покупка вопросов Вопросы = 11
    Case Else: GetCosts = "=-300"
  End Select
Exit Function
DateExit:
  ErrCollection 40, 1, 48, "для поставщика '" & Sh.Cells(PartRow, 5) & "' " _
    & IIf(Len(Sh.Cells(PartRow, 6)) > 0, "на " & Sh.Cells(PartRow, 6), _
    "в строке #" & PartRow): GetCosts = "=-100" ' EPN = 1
End Function

' Поиск строки SuppNumRow с данными о Поставщике на листе «Архив»
Public Sub GetSuppRow(ByRef Sh As Worksheet, ByVal PartRow As Long)
Dim SuppName As String, SuppRow As Integer, PartDate As Variant
  ''Stop
  SuppNumRow = 0: SuppRow = SuppRow + 1 ' Включаем счётчик строк
  SuppName = Sh.Cells(PartRow, 5)
  PartDate = Sh.Cells(PartRow, 6)
  ' ВАЖНО! Обновление списка с Индексами листов
  With App_Wb.Sheets(GetSheetList(Set_arName))
    cnfRenew = .Name
    'If IsEmpty(.Cells(SuppRow, 10)) Then SuppRow = SuppRow + 1 ' Костыль
    Do Until IsEmpty(.Cells(SuppRow, 10)) ' Счётчик строк ' Выполнять ДО
      If .Cells(SuppRow, 10) = SuppName And .Cells(SuppRow, 11) <= PartDate Then
        ' ВАЖНО! Если следующая «Дата актуальности» > PartDate, то закончить
        If .Cells(SuppRow, 10) <> .Cells(SuppRow + 1, 10) Or _
          (.Cells(SuppRow, 10) = .Cells(SuppRow + 1, 10) _
        And .Cells(SuppRow + 1, 11) > PartDate) Then _
          SuppNumRow = SuppRow: Exit Sub
      End If: SuppRow = SuppRow + 1
    Loop: If SuppNumRow = 0 And IsDate(PartDate) Then ErrCollection 30, 1, 16, _
      "для поставщика '" & SuppName & "' на " & PartDate ' EPN = 1
  End With
End Sub

Private Sub SendKeyEnter() ' Эмуляция нажатия клавиши «Enter»
  SendKeys "{ESC}", True: SendKeys "{ENTER}", False ' Костыль
End Sub

Private Sub SendKeysCtrlV() ' Эмуляция нажатия клавиш «Ctrl+V»
  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
End Sub