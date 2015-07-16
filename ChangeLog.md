### Условные обозначения

:exclamation: Новое, :star: Расширено, :star2: Оптимизировано, :fire: Исправлено, :grey_question: Тестовая версия

### r330

- :fire: Модуль ***Statistics.bas***
	- :exclamation: Процедура `SendKeysCtrlV`: перехват клавиш <kbd>Ctrl+V</kbd>
	- :star2: Процедура `Auto_Open`: проверка существования пути файла с ценами
	- :star: Процедура `SpecificationSheets`: небольшие изменения в форматировании ячеек
- :star: Модуль ***Frame.bas***
	- :exclamation: Процедура `SettingsStatistics`: сетевой путь в коллекцию с настройками
	- :exclamation: Процедура `ErrCollection`: сообщение о необходимости выбрать поставщика
- :fire: Класс ***cExcelEvents.cls***
	- :star: Процедура `Class_Initialize`: задаёт горячие клавиши процедурой `SendKeysCtrlV`
	- :star2: Процедура `App_WorkbookDeactivate`: копирование выделенного диапазона
	- :grey_question: Процедура `App_SheetSelectionChange`: в листах `SF_` и `SB_` изменили выпадающий список
- :star: Модуль ***AutoModuleRibbon.bas***
	- :exclamation: Процедура `GetVisibleMenu`: отображает рабочую вкладку меню только для текщей книги
	- :exclamation: Процедура `SetFilter`: управление кнопками фильтров на вкладке меню

### r320

- :exclamation: Модуль управления меню ***AutoModuleRibbon.bas***
- :grey_question: Файл ленточного меню ***customUI14.xml***
- :star: Колонки "*Дата материала*", "*Форма договора*", "*Кодекс*" на листах `SF_` и `SB_`
- :star2: Переставлены колонки "*Дата акта*" <=> "*Номер акта*", "*Дата договора*" <=> "*Номер договора*"

### r310

- :fire: Модуль ***Statistics.bas***
	- :star: Функция `CheckSupplier`: обновление массива **SuppDiff**
- :fire: Класс ***cExcelEvents.cls***
	- :star: Процедура `App_SheetActivate`: обновление массива **SuppDiff**
	- :star: Процедура `App_SheetSelectionChange`: для листа `SUPP_` создаём массив **SuppDiff**

### r300

- :grey_question: Шаблон книги ***blank r300.xlsx***: тестирование производится на первых 5 листах
- :exclamation: Основной модуль ***Statistics.bas***
	- :exclamation: Процедура `Auto_Open` (автозапуск): создаёт системную таблицу и загружает файл с ценами через `ADODB`
	- :grey_question: Процедура `SpecificationSheets`: восстанавливает форматирование таблиц
	- :exclamation: Процедура `RecordCells`: записывает на лист `ARCH_` данные о поставщике из массива **SuppDiff**
	- :exclamation: Функция `CheckSupplier`: проверяет изменения на листе `SUPP_` с массивом **SuppDiff**
	- :exclamation: Функция `ChangedBeforeSave`: для класса ***cExcelEvents.cls***
	- :exclamation: Процедура `ListCost`: создаёт список "*Категория цен*"
	- :exclamation: Функция `GetCosts`: возвращает цены на актуальную дату
	- :exclamation: Процедура `GetSuppRow`: поиск строки на листе `ARCH_`
	- :grey_question: Процедура `SendKeyEnter`: выполнение нажатия клавиши <kbd>Enter</kbd>
- :exclamation: Дополнительный модуль ***Frame.bas***
	- :exclamation: Свойство `GetUserName`: читает имя активного пользователя
	- :exclamation: Процедура `SettingsStatistics`: создаёт коллекцию с настройками
	- :exclamation: Функция `GetSheetList`: обновляет коллекцию с индексами листов и возвращает индекс листа по имени
	- :exclamation: Процедура `ProtectSheet`: защищает лист от изменений
	- :exclamation: Функция `UnprotectSheet`: снимает защиту с листа и возвращает сам объект
	- :exclamation: Процедура `SortSupplier`: выполняет сортировку по возрастанию по номеру колонки
	- :exclamation: Процедура `RemoveCollection`: удаляет все строки в коллекции
	- :exclamation: Функция `MultidimArr`: заполнение одномерного массива из двумерного
	- :exclamation: Процедура `ErrCollection`: выводит сообщение об ошибке по номеру и маркеру
- :exclamation: Класс событий книги ***cExcelEvents.cls***
	- :exclamation: Процедура `Class_Initialize`: объявляет приложение `App`, задаёт горячие клавиши процедурой `SendKeyEnter`
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

- :exclamation: Модуль ***DBCreateMDWSystem.bas***: создаёт системную таблицу ***System.mdw*** если не установлен **Access**

### r100

- Разработана структура файла с ценами ***Cost.accdb***