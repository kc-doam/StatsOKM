Attribute VB_Name = "Statistics"
Option Explicit
Option Base 1
'12345678901234567890123456789012345bopoh13@ya67890123456789012345678901234567890
Public Const Let_SuppList = "Список_Поставщиков", revFile As Integer = 350
' Внутреннее имя листа «Настройки», Внутреннее имя листа «Архив» и «Поставщики»
Public Const Set_cnfName = "CONF_", Set_arName = "ARCH_", Set_spName = "SUPP_"
Public Const Let_UrgencyTransList = "$Новое,$Наши вопросы,$Доработка,$Иное"
Public Const Let_ContrFormList = "станд.,опер."
Private Const Let_OrgBodyList = "Ф/Л,Ю/Л" ' Не менять!

' Рабочая книга, Рабочий лист, Имя рабочего листа/параметра
Public App_Wb As Workbook, cnfRenew As String, cstPath As String
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
Dim mdwPath As String, LastRow As Long
Dim Conn As Object, Rec As Object, Src As String
Const Let_mdwPath = "\Application Data\Microsoft\Access\System.mdw"
Const Let_cstPath = "\Архив\Cost.accdb" ' Цены
  SettingsStatistics Settings ' Загрузка настроек книги в коллекцию
  PartNumRow = ActiveCell.Row ' Номер строки партии материалов
  
  ' Проверка существования директории с настройками rev.330
  Set Conn = CreateObject("Scripting.FileSystemObject") ' fso
  cstPath = Settings("SetPath") & Let_cstPath
  If Not Conn.FileExists(cstPath) Then _
    ErrCollection 59, 1, 16, cstPath: Quit = True ' EPN = 1
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
    Loop: mdwPath = Environ("UserProfile") & Let_mdwPath ' Полный путь к файлу
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
    If Err Then ErrCollection Err.Number, 1, 16, cstPath: Quit = True ' EPN = 1
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
          & "Вопросы" & IIf(cnfRenew = "Ф/Л", ", [Оф письма], Кодекс", "") _
          & " FROM [" & SuppDiff(LBound(SuppDiff, 1), Counter) & "]" ' rev.340
        ' !Последний UNION ALL - НЕ ВКЛЮЧАТЬ!
        If Counter < UBound(SuppDiff, 2) Then Src = Src & " UNION ALL "
      Next Counter
      On Error Resume Next
        ' При открытии пустого объекта Recordset свойства BOF и EOF содержат True
        .Open Source:=Src, ActiveConnection:=Conn
        Cost.Add .GetRows(Rows:=-1), cnfRenew: .Close
      Debug.Print cnfRenew & " - Err #" & Err.Number; "; Fields "; .Fields.Count
        If Err.Number = 457 Then
          ErrCollection Err.Number, 1, 16, cnfRenew: Quit = True ' EPN = 1
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
  
  App_Wb.Sheets(Sh_List("SF_")).Activate ' ВАЖНО! Уйти с листа «Поставщики»
  For Each Rec In App_Wb.NameS ' Диапазон «Список_Поставщиков»
    If Rec.Name = Let_SuppList Then Rec.Delete
  Next Rec
  With ActiveWorkbook.NameS.Add(Name:=Let_SuppList, RefersTo:="=$E$1")
    .Comment = "Список листа Поставщики": .RefersTo = "=OFFSET('" _
      & App_Wb.Sheets(Sh_List(Set_spName)).Name & "'!$J$1,1,,COUNTA('" _
      & App_Wb.Sheets(Sh_List(Set_spName)).Name & "'!$J:$J)-1,)"
  End With: If CostChanged Then CostUpdate ' rev.340
  
  If Not App_Wb.ReadOnly Then
    For Each Conn In App_Wb.Sheets
      If Conn.CodeName = Set_cnfName Then UnprotectSheet Conn: _
        Conn.Range(Set_cnfName & "CostDate") = SuppDiff: SuppDiff = Empty
      ProtectSheet Conn ' rev.340
    Next Conn: cnfRenew = "": App_Wb.Save ' Если доступен, сохранить
  End If: cnfRenew = ActiveSheet.Name
End Sub

Public Function CostChanged() As Boolean ' rev.340
Dim Conn As Object, LastCostDate As Long
  Set Conn = CreateObject("Scripting.FileSystemObject") ' fso
  If Conn.FileExists(cstPath) Then LastCostDate = Mid(Log( _
    Conn.GetFile(cstPath).DateLastModified) - 10, 3, 8): Set Conn = Nothing
  If LastCostDate > 0 Then
    ' Проверка изменения файла с ценами
    If LastCostDate <> Settings("CostDate") Then
      If Len(cnfRenew) > 0 And Settings("CostDate") > 0 Then _
        ErrCollection 30, 1, 64 ' EPN = 1
      CostChanged = True
    End If: SuppDiff = LastCostDate
  End If
End Function

'''
Public Sub CostUpdate(Optional ByVal Supplier As String = "*") ' rev.340
Dim App_Sh As Worksheet, LastRow As Integer, Rec As Object
  Application.StatusBar = "Пожалуйста, подождите. Идёт обновление цен..."
  For Each App_Sh In App_Wb.Sheets ' Процедура пересчёта итоговых сумм
    If App_Sh.CodeName Like "[OQS]?_" Then
      With UnprotectSheet(App_Sh)
        '.Activate:
        LastRow = .UsedRange.Rows.Count ' Количество строк
        If LastRow > 1 Then ' Костыль (1 = должен быть Settings("head"))
        
          For Each Rec In .Range(.Cells(2, 44), .Cells(LastRow, 44))
            If .Cells(Rec.Row, 5) Like Supplier Then ' Not IsEmpty(.Cells(Rec.Row, 5))
              GetSuppRow App_Sh, Rec.Row ' Поиск строки SuppNumRow у поставщика
              SetFormula App_Sh, Rec.Row ' Создание формул для записи rev.350
              If .Cells(Rec.Row, 44).Value < 0 Then SuppDiff = 0 ' rev.340
            End If
          Next Rec
        End If
      End With: ProtectSheet App_Sh
    End If
  Next App_Sh: Application.StatusBar = False
End Sub

Public Sub SetFormula(ByRef Sh As Worksheet, ByVal TargetRow As Long) ' rev.350
Dim Arr As Variant
  Arr = Sh.Range(Sh.Cells(TargetRow, 1), Sh.Cells(TargetRow, 49)).Value
  Select Case Sh.CodeName
    Case "KF_", "BO_"
      Sh.Cells(TargetRow, 2).NumberFormat = "General" ' rev.340
      Sh.Cells(TargetRow, 6).NumberFormat = "m/d/yyyy"
      Arr(1, 1) = "=TEXT(R" & TargetRow & "C6,""ММММ.ГГ"")"
      Arr(1, 2) = "='" & cnfRenew & "'!R" & SuppNumRow & "C4"
      Arr(1, 3) = "='" & cnfRenew & "'!R" & SuppNumRow & "C5"
      Arr(1, 4) = "='" & cnfRenew & "'!R" & SuppNumRow & "C8"
      'Arr(1, 28) = "=SUM(R" & TargetRow & "C25:R" & TargetRow & "C27)"
      Arr(1, 32) = "=SUM(R" & TargetRow & "C29:R" & TargetRow & "C31)"
      ' НУМ Arr(1, 42) = "=SUM(R" & TargetRow & "C38:R" & TargetRow & "C41)"
      Arr(1, 44) = GetCosts(Sh, TargetRow)
      Arr(1, 46) = "=IF('" & cnfRenew & "'!R" & SuppNumRow & "C13=""НДС"",(R" & TargetRow & "C44+R" & TargetRow & "C45)*0.18,IF('" & cnfRenew & "'!R" & SuppNumRow & "C13=""УСН"",""без НДС"",""""))"
      Arr(1, 47) = "=SUM(R" & TargetRow & "C44:R" & TargetRow & "C46)"
      ' ЦЕНЫ БО
      'Sh.Cells(TargetRow, 57).FormulaR1C1 = "=SUM(R" & TargetRow & "C53:R" & TargetRow & "C56)"
      'Sh.Cells(TargetRow, 70).FormulaR1C1 = GetCosts(Sh, TargetRow)
      'Sh.Cells(TargetRow, 71).FormulaR1C1 = "=IF('" & cnfRenew & "'!R" & SuppNumRow & "C13=""НДС"",R" & TargetRow & "C70*0.18,IF('" & cnfRenew & "'!R" & SuppNumRow & "C13=""УСН"",""без НДС"",""""))"
      'Sh.Cells(TargetRow, 73).FormulaR1C1 = "=SUM(R" & TargetRow & "C70:R" & TargetRow & "C72)"
      'Sh.Cells(TargetRow, 74).FormulaR1C1 = "=R" & TargetRow & "C64+R" & TargetRow & "C73"
  End Select
  Sh.Range(Sh.Cells(TargetRow, 1), Sh.Cells(TargetRow, 49)).Value = Arr
  Sh.Cells(TargetRow, 2).NumberFormat = "@"
End Sub

' Установка рабочей конфигурации листов
Public Sub SpecificationSheets(ByVal SheetIndex As Byte)
Dim App_Sh As Worksheet, LastRow As Long, PreError As Variant
  On Error GoTo DataExit ' rev.350
    'Stop ' #3 Копирование заголовка, Установка условного форматирования
    For Each App_Sh In App_Wb.Sheets
      Application.ScreenUpdating = False ' ВЫКЛ Обновление экрана
      With UnprotectSheet(App_Wb.Sheets(App_Sh.Index))
        PreError = 0: .Activate ': UnprotectSheet App_Sh
        With ActiveWindow ' CTRL+HOME
          .ScrollRow = 1: .ScrollColumn = 1: .FreezePanes = False
        End With
        
        LastRow = App_Sh.UsedRange.Rows.Count ' Количество строк rev.330
        ' Очистка форматов, Очистка условного форматирования
        .Cells.ClearFormats: .Cells.FormatConditions.Delete
        ' Очистка группировки, Очистка проверки данных
        .Cells.ClearOutline: .Cells.Validation.Delete
        ' Очистка границы ячеек
        For Each PreError In Array(xlDiagonalDown, xlDiagonalUp, _
          xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, _
          xlInsideVertical, xlInsideHorizontal)
          .Cells.Borders(PreError).LineStyle = xlNone
        Next PreError
        
        Select Case .CodeName
          Case Set_spName, Set_arName ' Лист «Поставщики», «Архив»
            If .CodeName = Set_arName Then App_Wb.Sheets(Sh_List(Set_spName)) _
              .Range("A1:O1").Copy Destination:=.Range("A1")
              'SendKeys "^{HOME}", False ' rev.250 Фокус должен быть на MS Excel
              '.Cells(1, 1).AutoFilter
              
            If .AutoFilterMode Then .ShowAllData Else .Cells(1, 1).AutoFilter ' Автофильтр
            PreError = PreError + 1
            
            .Columns("C:I").Columns.Group: .Columns("F:G").Columns.Group
            If .CodeName = Set_spName Then
              .Columns("A:AB").Locked = False: .Rows("1:1").Locked = True
              .Columns("N:X").Columns.Group
            End If
            ' Форматирование колонок
            .Columns("D:D").NumberFormat = "@"
            .Columns("K:K").NumberFormat = "m/d/yyyy"
            If .CodeName = Set_spName Then
              .Range("A1:AB1").WrapText = True ' rev.340
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
            PreError = PreError + 1
            ' Условное форматирование
            If Val(Application.Version) >= 12 Then
              With .Range("A2:A55,D2:E55,L2:M55").FormatConditions _
                .Add(Type:=xlBlanksCondition)
                .Interior.ColorIndex = 3: .StopIfTrue = True
              End With
              With .Range("D2:D55").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=И($E2=""РИЦ"";$D2>""999"")")
                .Interior.ColorIndex = 3: .StopIfTrue = True ' rev.350
              End With
              With .Range("C2:C55").FormatConditions _
                .Add(Type:=xlBlanksCondition)
                .Interior.ColorIndex = 44: .StopIfTrue = True
              End With
              ' Поставщик (кратко)
              With .Range("J2:K55").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=И(НЕ(ЕПУСТО($J2));" _
                  & "ИЛИ(ЕПУСТО($A2);ЕПУСТО($L2);ЕПУСТО($M2)))")
                .Font.ColorIndex = 2: .Interior.ColorIndex = 9
                .StopIfTrue = True: .SetFirstPriority
              End With
              If .CodeName = Set_spName Then
                ' ИНН был перемещён ниже, а J55 сделать + 1
                With .Range("J2:J55").FormatConditions _
                  .Add(Type:=xlExpression, Formula1:="=И(НЕ(ЕПУСТО(СМЕЩ($J2;-1;" _
                    & "0)));ЕПУСТО($J2))")
                  .Interior.ColorIndex = 36: .StopIfTrue = True
                End With
                ' ИНН
                With .Range("W2:W55").FormatConditions _
                  .Add(Type:=xlExpression, Formula1:="=ИЛИ(ЕСЛИ(НЕ(ЕЧИСЛО($W2));" _
                    & "1;ЦЕЛОЕ($W2)<>$W2);ИЛИ(И($A2=""Ф/Л"";ДЛСТР($W2)<11));" _
                    & "И($A2=""Ю/Л"";ДЛСТР($W2)>10))")
                  .Interior.ColorIndex = 44: .StopIfTrue = True
                End With
                .Range("O2:AB55").FormatConditions.Add Type:=xlNoBlanksCondition
                With .Range("O2:O55,U2:U55,Z2:Z55").FormatConditions _
                  .Add(Type:=xlBlanksCondition)
                  .Interior.ColorIndex = 36: .StopIfTrue = True
                End With
                With .Range("Q2:Q55,T2:T55,V2:V55").FormatConditions _
                  .Add(Type:=xlExpression, Formula1:="=$A2=""Ф/Л""")
                  .Interior.ColorIndex = 36: .StopIfTrue = True
                End With
                With .Range("N2:N55,P2:P55,X2:X55").FormatConditions _
                  .Add(Type:=xlExpression, Formula1:="=$A2=""Ю/Л""")
                  .Interior.ColorIndex = 36: .StopIfTrue = True
                End With
              End If
              If .CodeName = Set_arName Then
                With .Range("A2:O55").FormatConditions _
                  .Add(Type:=xlExpression, _
                    Formula1:="=ЕПУСТО($J2)+МАКС(--($J2:$J9=$J2)*$K2:$K9)=$K2")
                  .Interior.ColorIndex = 43: .StopIfTrue = True
                  '.SetFirstPriority
                End With
                With .Range("A2:O55").FormatConditions _
                  .Add(Type:=xlExpression, _
                    Formula1:="=И($K2>СЕГОДНЯ()-90;НЕ(ЕПУСТО($J2)))")
                  .Interior.ColorIndex = 27: .StopIfTrue = True
                  '.SetFirstPriority
                End With
              End If
            End If
            PreError = PreError + 1
            ' Проверка ввода данных
            If .CodeName = Set_spName Then
              With .Range("A2:A55").Validation
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                  Formula1:=Let_OrgBodyList
                .ErrorTitle = "Вид лица"
                .ErrorMessage = "Необходимо выбрать значение из списка "
                .ShowError = True: .IgnoreBlank = True
              End With
              With .Range("D2:D55").Validation
                .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, _
                  Formula1:="=OR(AND($D2>=""001"",$D2<=""999"",LEN($D2)<4)," _
                  & "$D2=""станд."",$D2=""ДПР"")"
                .ErrorTitle = "Источник"
                .ErrorMessage = "Необходимо ввести 3-х значный номер РИЦа, " _
                  & "либо указать источник ""ДПР"" или ""станд."""
                .ShowError = True: .IgnoreBlank = True
              End With
              With .Range("E2:E55").Validation
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                  Formula1:="Коммерч.,Некоммерч.,Ведомство,РИЦ," _
                  & "Собств. источники КЦ"
                .ErrorTitle = "Тип организации"
                .ErrorMessage = "Необходимо выбрать значение из списка "
                .ShowError = True: .IgnoreBlank = True
              End With
              With .Range("K2:K55").Validation
                .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
                  Operator:=xlGreaterEqual, Formula1:=Replace( _
                  Settings("date0"), "#", "")
                .ErrorTitle = "Дата актуальности"
                .ErrorMessage = "Необходимо ввести дату не раньше " & .Formula1
                .ShowError = True: .IgnoreBlank = True
              End With
              With .Range("M2:M55").Validation
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                  Formula1:="НДС,УСН"
                .ErrorTitle = "НДС / УСН"
                .ErrorMessage = "Необходимо выбрать значение из списка "
              End With
              With .Range("Q2:Q55").Validation
                .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
                  Operator:=xlGreaterEqual, Formula1:=DateAdd( _
                  "m", -840, Replace(Settings("date0"), "#", ""))
                .ErrorTitle = "Дата рождения"
                .ErrorMessage = "Необходимо ввести дату не раньше " & .Formula1
                .ShowError = True: .IgnoreBlank = True
              End With
              With .Range("R2:R55").Validation
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                  Operator:=xlGreaterEqual, Formula1:=DatePart( _
                  "yyyy", Replace(Settings("date0"), "#", ""))
                .ErrorTitle = "Заявление о проф. вычете"
                .ErrorMessage = "Необходимо ввести год не меньше " & .Formula1
                .ShowError = True: .IgnoreBlank = True
              End With
              With .Range("V2:V55").Validation
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                  Operator:=xlBetween, Formula1:="100000000", _
                  Formula2:="99999999999"
                .ErrorTitle = "СНИЛС"
                .ErrorMessage = "Страховой номер в пенсионном фонде должен " _
                  & "содержать от 9 до 11 цифр "
                .ShowError = True: .IgnoreBlank = True
              End With
              With .Range("W2:W55").Validation ' ИНН
                .Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, _
                  Formula1:="=AND(OR(AND($A2=""Ф/Л"",LEN($W2)<13)," _
                  & "AND($A2=""Ю/Л"",LEN($W2)<11)),LEN($W2)>8)"
                .ErrorTitle = "ИНН"
                .ErrorMessage = "Идентификационный номер налогоплательщика " _
                  & "должен содержать: " & vbCrLf & vbTab & "для Ф/Л  от 11 " _
                  & "до 12 цифр" & vbCrLf & vbTab & "для Ю/Л  от 9 до 10 цифр"
                .ShowError = True: .IgnoreBlank = True
              End With
              With .Range("X2:X55").Validation
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
            If .CodeName = Set_spName Then ' Закрепление области rev.350
              .Range("K2").Select: ActiveWindow.FreezePanes = True
            End If
          Case "SF_", "SB_"
          .Tab.ColorIndex = 24: LastRow = LastRow + 99
              'SendKeys "^{HOME}", False ' rev.250 Фокус должен быть на MS Excel
              '.Cells(1, 1).AutoFilter
              
            If .AutoFilterMode Then .ShowAllData Else .Cells(1, 1).AutoFilter ' Автофильтр
            PreError = PreError + 1
            
            .Range("A1:AW1").WrapText = True ' rev.340
            .Columns("E:I").Locked = False: .Columns("K:R").Locked = False
            '.Columns("T:AA").Locked = False ' rev.330
            .Columns("T:AE").Locked = False: .Columns("AG:AQ").Locked = False ' rev.350
            .Columns("AS").Locked = False
            .Columns("AV:AW").Locked = False: .Rows("1:1").Locked = True
            .Columns("B:D").Columns.Group: .Columns("I:S").Columns.Group ' rev.340
            .Columns("Y:AA").Columns.Group: .Columns("AC:AE").Columns.Group
            '.Columns("AG:AQ").Columns.Group: .Columns("AL:AO").Columns.Group
            .Columns("AG:AO").Columns.Group: .Columns("AL:AN").Columns.Group ' rev.350
            
            ' Форматирование колонок
            .Columns("A:D").NumberFormat = "General"
            .Columns("B:B").NumberFormat = "@" ' rev.340
            With .Columns("X:AU") ' rev.350
              .NumberFormat = "#,##0"
              .HorizontalAlignment = xlRight: .IndentLevel = 1
            End With: .Rows("1:1").HorizontalAlignment = xlGeneral
            With .Columns("F:X") ' rev.350
              .NumberFormat = "m/d/yyyy"
              .HorizontalAlignment = xlGeneral
            End With: .Range("P:P,R:R,S:S").NumberFormat = "@" ' rev.350
            ' Границы таблицы rev.350
            With .Range("D:D,N:N,AA:AA,AE:AE,AI:AI,AK:AK,AP:AP") _
              .Borders(xlEdgeRight)
              .LineStyle = xlContinuous: .Weight = xlThin
            End With
            With .Range("X:X,AB:AB,AF:AF,AQ:AQ,AR:AR,AU:AU") _
              .Borders(xlEdgeRight)
              .LineStyle = xlContinuous: .Weight = xlMedium
            End With
            ' Сортировка
            SortSupplier App_Sh, 6, 7
            PreError = PreError + 1
            ' Условное форматирование
            If Val(Application.Version) >= 12 Then
              ' Ошибка «Дата актуальности» или отсуствие цены в массиве Cost
              With .Range("A:AW").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=$AR1<0")
                .Font.ColorIndex = 2: .Interior.ColorIndex = 9
                .StopIfTrue = True: .SetFirstPriority
              End With
              With .Range("F2:O" & LastRow & ",Q2:Q" & LastRow & ",T2:X" & LastRow).FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=И(ИЛИ(F2<" & Replace( _
                  Settings("date0"), "#", "") & ";F2>СЕГОДНЯ());F2<>"""";F2<>""не оплач."")")
                .Interior.ColorIndex = 3: .StopIfTrue = True
              End With
              With .Range("F2:G" & LastRow).FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=СЧЁТЕСЛИМН($E2:$E" & LastRow & ";$E2;$F2:$F" & LastRow & ";$F2;$G2:$G" & LastRow & ";$G2)>1")
                .Interior.ColorIndex = 3: .StopIfTrue = True
              End With
              With .Range("G2:G" & LastRow).FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=СЧЁТЕСЛИМН($E2:$E" & LastRow & ";$E2;$G2:$G" & LastRow & ";$G2)>1")
                .Interior.ColorIndex = 45: .StopIfTrue = True
              End With
              
              With .Range("K:N,X:X").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=$X1=""не оплач.""")
                .Interior.ColorIndex = 48: .StopIfTrue = True
              End With
              With .Range("K:K").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=И($J1<СЕГОДНЯ()-30;$J1<>"""";$K1="""")")
                .Interior.ColorIndex = 38: .StopIfTrue = True
              End With
              With .Range("L:L").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=И($K1<СЕГОДНЯ()-9;$K1<>"""";$L1="""")")
                .Interior.ColorIndex = 38: .StopIfTrue = True
              End With
              With .Range("V:V").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=И($U1<СЕГОДНЯ()-14;$U1<>"""";$V1="""")")
                .Interior.ColorIndex = 38: .StopIfTrue = True
              End With
              With .Range("O:T").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=$X1<>""""")
                .StopIfTrue = True ' Без цвета
              End With
              With .Range("O:T").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=$U1<>""""")
                .Interior.ColorIndex = 43: .StopIfTrue = True
              End With
              With .Range("O:T").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=$L1<>""""")
                .Interior.ColorIndex = 44: .StopIfTrue = True
              End With
              
              With .Range("Y:AA").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=И(ДЛСТР(СЦЕПИТЬ($Y1;$Z1;$AA1))<1;СУММ($Y1:$AA1)<>$AB1)")
                .Interior.ColorIndex = 44: .StopIfTrue = True
              End With
              With .Range("Y2:AB" & LastRow).FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=И(СУММ($Y2:$AA2)<>$AB2;$AF2>0)")
                .Interior.ColorIndex = 3: .StopIfTrue = True
              End With
              With .Range("AL2:AP" & LastRow).FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=СУММ($AL2:$AO2)<>$AP2")
                .Interior.ColorIndex = 3: .StopIfTrue = True
              End With
              With .Range("AG2:AK" & LastRow & ",AQ2:AR" & LastRow).FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=ИЛИ(СУММ($AG2:$AK2)<>ЕСЛИ($AQ2>0;0;СУММ(ЕСЛИ(ЕЧИСЛО(ПОИСК(""вопр"";$AV2));$AJ2:$AK2;$AG2:$AI2)));$AF2<>ЕСЛИ($AQ2>0;$AQ2;СУММ($AG2:$AK2)))")
                .Interior.ColorIndex = 3: .StopIfTrue = True
              End With
              With .Range("M:M").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=И($L1>0;$M1="""")")
                .Interior.ColorIndex = 38: .StopIfTrue = True
              End With
              'With .Range("O:T").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=И($M1>0;$Q1>0;$N1="""")")
              '  .Interior.ColorIndex = 38: .StopIfTrue = True
              'End With
              
              ' Срочно Наши вопросы
              With .Range("E:N,X:AF,AJ:AK,AR:AW").FormatConditions _
                .Add(Type:=xlExpression, Formula1:="=$AV1=""Срочно Наши вопросы""")
                .Interior.ColorIndex = 37: .StopIfTrue = False
              End With
            End If
            PreError = PreError + 1
            ' Проверка ввода данных
            
            'With
            '  .NumberFormat = "#,##0"
            '  .HorizontalAlignment = xlRight
            '  .IndentLevel = 1
            'End With
            
            ' Закрепление области
            .Range("G2").Select: ActiveWindow.FreezePanes = True
            With Selection ' Выделяем ячейку в последней строке rev.350
              If LastRow > 2 ^ 7 Then .Offset(, -2).End(xlDown).Select
            End With
        End Select: .Outline.ShowLevels ColumnLevels:=1: ProtectSheet App_Sh
      End With
    Next App_Sh
    'On Error Resume Next
    '  App_Wb.Sheets(SheetIndex).Select
    '  If Err Then App_Wb.Sheets(Sh_List("SF_")).Select
    App_Wb.Sheets(Sh_List("SF_")).Select
    'On Error GoTo 0
    Application.ScreenUpdating = True ' ВКЛ Обновление экрана
    'SendKeys "{NUMLOCK}", True ' Костыль v2.5 Фокус должен быть на MS Excel
Exit Sub
DataExit:
  ErrCollection Err.Number + PreError, 3, 48, App_Sh.Name ' EPN = 3
End Sub

' Запись в «Архив» данных о поставщике из массива SuppDiff
Public Sub RecordCells(ByVal NewSupplier As Boolean)
Dim i As Integer: Counter = 0: i = 2
  'Stop ''Stop - i становится в конец
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
            If Err Then ErrCollection Err.Number, 1, 16: Quit = True ' EPN = 1
            'Debug.Print "RecordCells Err #" & Err.Number: If Err Then Exit Sub
          Loop: If Counter <> 0 Then i = Counter
          
          .Cells(i, 1).Resize(, UBound(SuppDiff)) = SuppDiff
          If IsDate(SuppDiff(11)) Then .Cells(i, 11) = CDate(SuppDiff(11))
          
          SortSupplier App_Wb.Sheets(Sh_List(Set_arName)), 10, 11
          CostUpdate SuppDiff(10) ' rev.340
          
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

''' ???
Public Function GetCosts(ByRef Sh As Worksheet, ByVal PartRow As Long) As String
Dim SuppCost As Variant, PartDate As Variant, OrgBody As String ' rev.340
  On Error GoTo DateExit ' rev.340 If IsEmpty(Cost) Then GoTo DateExit
    'Stop '' STOP
    PartDate = Sh.Cells(PartRow, 6)
    With App_Wb.Sheets(Sh_List(Set_arName))
      OrgBody = .Cells(SuppNumRow, 1) ' rev.340
      For Counter = LBound(Cost(OrgBody), 2) To UBound(Cost(OrgBody), 2)
      
      Debug.Print Cost(OrgBody)(0, Counter); "="; .Cells(SuppNumRow, 12)
      Debug.Print Cost(OrgBody)(1, Counter); "<="; PartDate
      ''Stop
        
        If Cost(OrgBody)(0, Counter) = .Cells(SuppNumRow, 12) _
        And Cost(OrgBody)(1, Counter) <= PartDate Then
          ' ВАЖНО! Если следующее поле цены «Актуально» > «Даты актуальности»
          If Cost(OrgBody)(0, Counter) <> Cost(OrgBody)(0, Counter + 1) _
          Or (Cost(OrgBody)(0, Counter) = Cost(OrgBody)(0, Counter + 1) _
          And Cost(OrgBody)(1, Counter + 1) > PartDate) Then _
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
              'Case 14-: GetCosts = GetCosts & "+RC[-2]*" & PartRow ' Бухонлайн Ф/Л
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
  'Debug.Print "SuppNumRow ="; SuppNumRow
  If ActiveCell.Column = 6 And IsEmpty(ActiveCell) Then
    GetCosts = "=0"
  Else
    If IsEmpty(SuppDiff) Then ErrCollection 40, 1, 48, "для поставщика '" _
      & Sh.Cells(PartRow, 5) & "' " & IIf(Len(PartDate) > 0, _
      "на " & PartDate, "в строке #" & PartRow) ' EPN = 1
    GetCosts = IIf(IsEmpty(Sh.Cells(PartRow, 5)), "=-100", "=-200")
    ' Ошибка, поэтому обновить суммы при открытии статистики
    With UnprotectSheet(App_Wb.Sheets(Sh_List(Set_cnfName)))
      .Range(Set_cnfName & "CostDate") = 0
    End With: ProtectSheet App_Wb.Sheets(Sh_List(Set_cnfName))
  End If
End Function

' Поиск строки SuppNumRow с данными о Поставщике на листе «Архив» rev.340
Public Sub GetSuppRow(ByRef Sh As Worksheet, ByVal PartRow As Long)
  SuppNumRow = 0: Counter = 1 + 1 ' Счётчик строк листа «Архив»
  ' ВАЖНО! Обновление списка с Индексами листов
  With App_Wb.Sheets(GetSheetList(Set_arName))
    cnfRenew = .Name ' ВАЖНО! Передаём имя листа
    Do Until IsEmpty(.Cells(Counter, 10)) ' Счётчик строк ' Выполнять ДО
      If .Cells(Counter, 10) = Sh.Cells(PartRow, 5) _
      And .Cells(Counter, 11) <= Sh.Cells(PartRow, 6) Then
        ' ВАЖНО! Если следующая «Дата актуальности» > «Дата поступления», то
        If .Cells(Counter, 10) <> .Cells(Counter + 1, 10) Or _
          (.Cells(Counter, 10) = .Cells(Counter + 1, 10) _
        And .Cells(Counter + 1, 11) > Sh.Cells(PartRow, 6)) Then _
          SuppNumRow = Counter: Exit Sub
      End If: Counter = Counter + 1
    Loop
  End With
End Sub

Private Sub SendKeyEnter() ' Эмуляция нажатия клавиши «Enter»
  SendKeys "{ESC}", True: SendKeys "{ENTER}", False ' Костыль
End Sub

Private Sub SendKeysCtrlV() ' Эмуляция нажатия клавиш «Ctrl+V»
  Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone
End Sub
