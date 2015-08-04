### Условные обозначения

:exclamation: Новое, :star: Расширено, :star2: Оптимизировано, :fire: Исправлено, :grey_question: Тестовая версия

### r360

:star2: Переставлены колонки "*Дата актуальности*" <= перемещена перед "*Категория цены*"

:star2: В список "*Тип организации*" добавлена позиция "*Ведомство (без подп.)*"

:fire: Модуль ***Statistics.bas***
- :star: Процедура `SpecificationSheets`: переменной `LastRow` присваивается значение последней строки на листе; добавлена проверка дат на листах `SF_` и `SB_`; небольшие изменения в форматировании ячеек
- :fire: Процедура `RecordCells`: изменения в массиве **SuppDiff** и параметрах процедуры `SortSupplier` из-за перестановки колонок
- :fire: Функция `CheckSupplier`: изменения в массиве **SuppDiff** из-за перестановки колонок
- :fire: Функция `ChangedBeforeSave`: изменения в массиве **SuppDiff** из-за перестановки колонок
- :fire: Процедура `ListCost`: небольшие изменения из-за перестановки колонок (аналогично процедура `SuppNumRow`)
- :fire: Функция `GetCosts`: небольшие изменения из-за перестановки колонок
- :star2: Процедура `SendKeysCtrlV`: проверка наличия неформатированного текста в буфере

:exclamation: Модуль ***Frame.bas***
- :exclamation: Свойство `Quit`: отображать *[Только для чтения]* в заголовке
- :exclamation: Процедура `ErrCollection`: сообщение об ошибке в формуле условного форматирования

:fire: Класс событий книги ***cExcelEvents.cls***
- :grey_question: Процедура `App_SheetBeforeRightClick`: небольшие изменения из-за перестановки колонок (аналогично процедура `App_SheetSelectionChange`)
- :star: Процедура `App_SheetChange`: удаление всех символов, кроме натуральных чисел; автопростановка "*не оплач.*" в колонке "*Дата перечислений*"

### r350

:fire: Модуль ***Statistics.bas***
- :exclamation: Функция `SetFormula`: создание формул в строке с поставщиком на листах `SF_` и `SB_`
- :star2: Процедура `SpecificationSheets`: выполнение функции `ErrCollection` в случае ошибки; добавление очистки границ ячеек; изменения в форматировании и группировки ячеек, перемещении курсора на последнюю строку

:exclamation: Модуль ***Frame.bas***
- :exclamation: Процедура `ErrCollection`: сообщение о невозможности создания условного форматирования

:fire: Класс событий книги ***cExcelEvents.cls***
- :grey_question: Процедура `App_SheetChange`: *отдельно для каждого листа*
- :star2: Процедура `App_SheetSelectionChange`: выполнение функции `SetFormula`; добавление выпадающего списка

### r340

:fire: Модуль ***Statistics.bas***
- :exclamation: Функция `CostChanged`: проверка изменения файла с ценами
- :exclamation: Функция `CostUpdate`: пересчёт формул для поставщика
- :star: Процедура `Auto_Open`: цены "*Бухонлайн*" объединены с ценами "*Кодекс*"; добавление свойства `Quit`
- :star: Процедура `SpecificationSheets`: небольшие изменения в форматировании ячеек
- :star2: Процедура `RecordCells`: выполнение функции `CostUpdate`; добавление свойства `Quit`
- :grey_question: Функция `GetCosts`: переменная `OrgBody` вынесена из параметров функции в локальную
- :star2: Процедура `GetSuppRow`: удалена проверка с сообщением об ошибке

:exclamation: Модуль ***Frame.bas***
- :exclamation: Свойство `Quit`: изменения параметры вставки для текущей книги
- :star2: Функция `GetSheetList`: добавление свойства `Quit`
- :star2: Процедура `ErrCollection`: сообщение о изменении в файле с ценами; добавление свойства `Quit`

:fire: Класс ***cExcelEvents.cls***
- :star: Процедура `App_SheetActivate`: на листах кроме `SUPP_` и `ARCH_` присвоить `PartNumRow` 
- :star: Процедура `App_SheetSelectionChange`: на листах `SF_` и `SB_` изменили условия выпадающих списоков
- :star: Процедура `App_SheetSelectionChange`: выполнение функции `CostChanged`
- :star2: Процедура `App_WorkbookActivate`: добавление свойства `Quit`
- :star2: Процедура `App_WorkbookBeforeClose`: добавление свойства `Quit`
- :star2: Процедура `App_WorkbookBeforeSave`: добавление свойства `Quit`
- :star2: Процедура `App_WorkbookDeactivate`: добавление свойства `Quit`

### r330

:fire: Модуль ***Statistics.bas***
- :exclamation: Процедура `SendKeysCtrlV`: перехват клавиш <kbd>Ctrl+V</kbd>
- :star: Процедура `SpecificationSheets`: небольшие изменения в форматировании ячеек
- :star2: Процедура `Auto_Open`: проверка существования пути файла с ценами

:fire: Модуль ***Frame.bas***
- :exclamation: Процедура `SettingsStatistics`: сетевой путь в коллекцию с настройками
- :star2: Процедура `ErrCollection`: сообщение о необходимости выбрать поставщика

:fire: Класс ***cExcelEvents.cls***
- :star: Процедура `Class_Initialize`: задаёт горячие клавиши процедурой `SendKeysCtrlV`
- :star2: Процедура `App_WorkbookDeactivate`: копирование выделенного диапазона
- :grey_question: Процедура `App_SheetSelectionChange`: на листах `SF_` и `SB_` изменили выпадающий список

:exclamation: Модуль ***AutoModuleRibbon.bas***
- :exclamation: Процедура `GetVisibleMenu`: отображает рабочую вкладку меню только для текущей книги
- :exclamation: Процедура `SetFilter`: управление кнопками фильтров на вкладке меню; процедура `AddFilter` удалена

### r320

:exclamation: Модуль управления меню ***AutoModuleRibbon.bas***

:grey_question: Файл ленточного меню ***customUI14.xml***

:star: Колонки "*Дата материала*", "*Форма договора*", "*Кодекс*" на листах `SF_` и `SB_`

:star2: Переставлены колонки "*Дата акта*" <=> "*Номер акта*", "*Дата договора*" <=> "*Номер договора*"

### r310

:fire: Модуль ***Statistics.bas***
- :star: Функция `CheckSupplier`: обновление массива **SuppDiff**

:fire: Класс ***cExcelEvents.cls***
- :star: Процедура `App_SheetActivate`: обновление массива **SuppDiff**
- :star: Процедура `App_SheetSelectionChange`: для листа `SUPP_` создаём массив **SuppDiff**

### r300

:grey_question: Шаблон книги ***blank r300.xlsx***: тестирование производится на первых 5 листах

:exclamation: Основной модуль ***Statistics.bas***
- :exclamation: Процедура `Auto_Open` (автозапуск): создаёт системную таблицу и загружает файл с ценами через `ADODB`
- :grey_question: Процедура `SpecificationSheets`: восстанавливает форматирование таблиц
- :exclamation: Процедура `RecordCells`: записывает на лист `ARCH_` данные о поставщике из массива **SuppDiff**
- :exclamation: Функция `CheckSupplier`: проверяет изменения на листе `SUPP_` с массивом **SuppDiff**
- :exclamation: Функция `ChangedBeforeSave`: для класса ***cExcelEvents.cls***
- :exclamation: Процедура `ListCost`: создаёт список "*Категория цен*"
- :exclamation: Функция `GetCosts`: возвращает цены на актуальную дату
- :exclamation: Процедура `GetSuppRow`: поиск строки на листе `ARCH_`
- :grey_question: Процедура `SendKeyEnter`: выполнение нажатия клавиши <kbd>Enter</kbd>

:exclamation: Дополнительный модуль ***Frame.bas***
- :exclamation: Свойство `GetUserName`: читает имя активного пользователя
- :exclamation: Процедура `SettingsStatistics`: создаёт коллекцию с настройками
- :exclamation: Функция `GetSheetList`: обновляет коллекцию с индексами листов и возвращает индекс листа по имени
- :exclamation: Процедура `ProtectSheet`: защищает лист от изменений
- :exclamation: Функция `UnprotectSheet`: снимает защиту с листа и возвращает сам объект
- :exclamation: Процедура `SortSupplier`: выполняет сортировку по возрастанию по номеру колонки
- :exclamation: Процедура `RemoveCollection`: удаляет все строки в коллекции
- :exclamation: Функция `MultidimArr`: заполнение одномерного массива из двумерного
- :exclamation: Процедура `ErrCollection`: выводит сообщение об ошибке по номеру и маркеру

:exclamation: Класс событий книги ***cExcelEvents.cls***
- :exclamation: Процедура `Class_Initialize`: объявляет приложение `App`; задаёт горячие клавиши процедурой `SendKeyEnter`
- :exclamation: Процедура `App_WorkbookOpen`: не используется
- :exclamation: Процедура `App_WorkbookActivate`: направление перемещения курсора "*вправо*"
- :exclamation: Процедура `App_WorkbookBeforeClose`: выполнение процедуры `RecordCells`
- :exclamation: Процедура `App_WorkbookBeforeSave`: выполнение процедуры `RecordCells` и `SpecificationSheets`
- :exclamation: Процедура `App_WorkbookDeactivate`: направление перемещения курсора "*вниз*" и выполнение процедуры `RecordCells`
- :exclamation: Процедура `App_SheetDeactivate`: *отдельно для каждого листа*
- :exclamation: Процедура `App_SheetActivate`: на лист `ARCH_`
- :grey_question: Процедура `App_SheetBeforeDoubleClick`: переключение с листа `SUPP_` на лист `ARCH_`
- :grey_question: Процедура `App_SheetBeforeRightClick`: запрет удаления строки
- :exclamation: Процедура `App_SheetSelectionChange`: *отдельно для каждого листа*

### r200

:exclamation: Модуль ***DBCreateMDWSystem.bas***: создаёт системную таблицу ***System.mdw*** если не установлен **Access**

### r100

:grey_question: Разработана структура файла с ценами ***Cost.accdb***
