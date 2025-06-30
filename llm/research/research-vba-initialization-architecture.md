Отлично. Я начну с создания первого документа — «Initialization & Entry-point». Он будет включать структурированный обзор механизмов и шаблонов инициализации в VBA-приложениях Excel, включая точку входа, загрузку компонентов, безопасную реализацию Singleton и другие ключевые аспекты архитектуры и запуска.

Я уведомлю вас, когда первый документ будет готов для просмотра.


# Initialization & Entry-point

Первый раздел этого руководства посвящён организации **точки входа** и **инициализации** Excel VBA-приложения. Здесь мы рассмотрим, как задать единый порядок запуска и настроить окружение, а также как централизованно загружать зависимости. Каждый подраздел состоит из: **Проблемы** → **Причины (VBA)** → **Ошибка (пример)** → **Решение (шаблон/приём)**, с наглядными примерами «Bad» → «Good».

## 1. Точка входа и начальная инициализация приложения

### 1.1 Нечёткая точка входа VBA‑макросов

**Проблема:** У Excel нет встроенного явного «Main». Код запускается по событиям (кнопки, события книги) и без централизованного порядка. Это ведёт к разбросу макросов и трудностям управления последовательностью инициализации.
**Причина:** VBA-макросы можно вызывать из разных мест (Auto\_Open, Workbook\_Open, кнопки на листах и т.д.), и без дисциплины их порядок непредсказуем. В частности, использование **Sub Auto\_Open()** считается устаревшим способом запуска макроса при открытии.
**Ошибка:** Код инициализации разбросан по модулям/формам. Например, используют `Auto_Open` или макросы привязанные к кнопкам на разных листах, нет единого порядка. Это приводит к тому, что некоторые процедуры могут не вызвать другие необходимые, а состояние приложения на старте остаётся неопределённым.
**Решение:** Определите одну **точку входа** – например, обработчик события `Workbook_Open` в модуле `ThisWorkbook`, который вызывает процедуру инициализации. Всегда включайте `Option Explicit` для обнаружения ошибок на этапе компиляции. Обновите `Application.EnableEvents`, `ScreenUpdating` и другие свойства вокруг вызова инициализации.

**Bad:** Разрозненный Auto\_Open в стандартном модуле.

```vba
'***** BAD: код запускается через Auto_Open (устарело)
Public Sub Auto_Open()
    ' Установить значение в ячейку текущего активного листа
    ActiveSheet.Range("A1").Value = "Start"
End Sub

Public Sub DoWork()
    ' Пример задачи без гарантии, что инициализация выполнена
    MsgBox "Doing work..."
End Sub
```

**Good:** Единая точка входа – событие `Workbook_Open` в `ThisWorkbook`, вызывающее процедуру инициализации.

```vba
Option Explicit

'***** GOOD: в модуле ThisWorkbook
Private Sub Workbook_Open()
    ' Заблокировать события и анимацию на время инициализации
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Call AppInitialize  ' центральная процедура инициализации

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

'***** GOOD: в стандартном модуле
Public Sub AppInitialize()
    Dim wb As Workbook
    Set wb = ThisWorkbook  ' явно заданная книга для инициализации

    Dim ws As Worksheet
    Set ws = wb.Worksheets("Data")
    ws.Range("A1").Value = "Initialized at " & Now  ' пример инициализации

    ' Дополнительная начальная настройка (напр., режим вычислений):
    Application.Calculation = xlManual  ' отключаем автоматические пересчёты
End Sub
```

> *На будущее:* всегда централизуйте вызов начальных процедур. Используйте `Workbook_Open` для автоматического запуска кода при открытии книги. Всегда `Option Explicit` и обрабатывайте флаги `EnableEvents/ScreenUpdating` вокруг инициализации.

**Summary:** При старте приложения используйте событие `Workbook_Open` (в `ThisWorkbook`) как единую точку входа. Отключайте ненужные события и обновление экрана во время инициализации.
**Golden Rules:**

* **Всегда** включайте `Option Explicit` во всех модулях.
* **Единый запуск:** всю первичную инициализацию вызывайте из `Workbook_Open`.
* **Контроль окружения:** в начале отключайте `ScreenUpdating`, `EnableEvents`; в конце – включайте обратно.
* **Явные объекты:** пользуйтесь `ThisWorkbook`, `Worksheet` и т.д., а не `ActiveSheet`, чтобы быть уверенным, что инициализируется нужная книга/лист.

**Masterclass:** демонстрирует полную последовательную загрузку приложения при открытии книги. Кладёт данные в лист, меняет параметры Excel и восстанавливает состояние.

```vba
Option Explicit

' ThisWorkbook:
Private Sub Workbook_Open()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    MainInitialize  ' главный метод инициализации

    Application.Calculation = xlAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Стандартный модуль:
Public Sub MainInitialize()
    Dim wb As Workbook
    Set wb = ThisWorkbook  ' чётко указываем нужную книгу

    ' Пример: инициализация данных на листах
    Dim wsData As Worksheet
    Set wsData = wb.Worksheets("Data")
    wsData.Range("A1").Value = "App started at " & Now

    Dim wsCfg As Worksheet
    Set wsCfg = wb.Worksheets("Config")
    wsCfg.Range("B1").Value = "Mode: Initialization complete"

    ' Другие действия...
End Sub
```

## 2. Организация загрузки зависимостей и объектов

### 2.1 Неявные глобальные зависимости (ActiveWorkbook/ActiveSheet)

**Проблема:** Процедуры оперируют «скрытыми» объектами VBA (как `ActiveWorkbook`, `ActiveSheet`, `Selection`), что делает код зависимым от текущего состояния Excel.
**Причина:** Если не указать явно, VBA обращается к глобальной среде (например, `Worksheets` равнозначен `Application.ActiveWorkbook.Worksheets`). При этом «Active» может быть любым — пользователь мог переключиться между книгами.
**Ошибка:** Процедуры используют `Worksheets.Add`, `Selection` или `ActiveWorkbook` без явной ссылки. Например, создаётся лист в неожиданной книге или форматируется выделение, не связанное с конкретной книгой.
**Решение:** Всегда передавайте в процедуры нужные объекты (Workbook, Worksheet, Range) параметрами, или используйте явно `ThisWorkbook.Worksheets(…)`. Это называется *внедрением зависимости* (Dependency Injection). Вместо `Worksheets.Add` делайте `wb.Worksheets.Add`, где `wb` — объект книги, переданный в процедуру.

**Bad:** Используется `Worksheets` без ссылки (меняет активную книгу):

```vba
Public Sub CreateReport()
    Dim ws As Worksheet
    Set ws = Worksheets.Add  ' Будет добавлять лист в ActiveWorkbook
    ws.Name = "Report"
End Sub
```

**Good:** Передаём книгу как параметр:

```vba
Public Sub CreateReport(ByVal wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add  ' Явно в заданной книге
    ws.Name = "Report"
End Sub

' В точке вызова:
Call CreateReport(ThisWorkbook)
```

### 2.2 Централизованная инициализация общих объектов

**Проблема:** Один и тот же объект (например, словарь, соединение с БД, фабричный класс) создаётся в разных местах, что дублирует код и расходует ресурсы.
**Причина:** Без выделенного места инициализации каждая процедура «по надобности» делает `New` или `CreateObject`, даже если нужен один экземпляр.
**Ошибка:** Код многократно создаёт объекты. Например, в каждом методе `Dim dict As New Scripting.Dictionary`, приводя к многократному заполнению и разным состояниям.
**Решение:** Создайте **единую точку инициализации** или «фабрику» для общих объектов. Инициализируйте объект один раз и переиспользуйте. Например, заведите `Private dictCache As Object` и функцию `GetDict()`, которая при первом вызове создаст экземпляр, а при последующих — вернёт уже созданный.

**Bad:** В каждой процедуре создаётся новый словарь:

```vba
Public Sub InitData()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Key", "Value"
    ' ... повторная инициализация каждый раз
End Sub
```

**Good:** Инициализируем один раз и храним в модуле:

```vba
Option Explicit

Private dictData As Object

Public Function GetDataDict() As Object
    If dictData Is Nothing Then
        Set dictData = CreateObject("Scripting.Dictionary")
        dictData.Add "Key", "Value"
    End If
    Set GetDataDict = dictData
End Function

Public Sub UseData()
    Dim d As Object
    Set d = GetDataDict()  ' всегда один и тот же экземпляр
    MsgBox d("Key")
End Sub
```

> **Совет:** Подобный приём – простой вариант шаблона Factory/Singleton. Он устраняет дублирование создания и гарантирует единый объект во всём приложении.

**Summary:** Избегайте «магических» глобальных обращений (`ActiveWorkbook`, `Selection`). Внедряйте необходимые объекты через параметры или централизованные функции. При общих объектах (напр., конфигурация, кэш) создавайте их один раз через приватную переменную в модуле.
**Golden Rules:**

* **Передавайте объекты параметрами** вместо использования `ActiveX`: все процедуры «знают» свои зависимости.
* **Централизуйте создание:** для общих объектов используйте одну функцию или свойство.
* **Избегайте дублирования:** если объект нужен в нескольких местах, храните его единственный экземпляр в модуле.
* **Читаемость:** явно объявляйте типы и используйте Option Explicit, чтобы не полагаться на неявное преобразование.

**Masterclass:** показывает загружаемый синглтон-объект (словарь конфигураций) и использование его в разных частях приложения.

```vba
Option Explicit

' Модуль ConfigManager: обеспечивает доступ к общим настройкам
Private cfg As Object

Public Function GetConfig() As Object
    If cfg Is Nothing Then
        Set cfg = CreateObject("Scripting.Dictionary")
        cfg.Add "Mode", "Test"
        cfg.Add "Version", "1.0"
    End If
    Set GetConfig = cfg
End Function

' Пример использования из разных мест:
Public Sub ShowConfig()
    Dim config As Object
    Set config = GetConfig()
    MsgBox "Mode=" & config("Mode")
End Sub

Public Sub UpdateConfig()
    Dim config As Object
    Set config = GetConfig()
    config("Mode") = "Production"  ' Меняем единственный экземпляр
End Sub
```

## 3. Безопасная реализация Singleton

**Проблема:** Для некоторых классов нужен ровно один экземпляр (конфигуратор, логгер, менеджер ресурсов). VBA не поддерживает приватный конструктор, поэтому легко создать несколько экземпляров случайно.
**Причина:** Класс VBA можно инстанцировать оператором `New` в любом месте. Без контроля можно получить несколько инстансов, что разрушит идею единственного хранилища данных.
**Ошибка:** Объект-одиночка объявляют просто `Public MyObj As New Class`, либо не освобождают. Например, каждый модуль делает `Dim log As New Logger` – получается несколько разных логгеров.
**Решение:** Реализуйте шаблон **Singleton** вручную через модуль. В модуле заведите приватную переменную (`Private objShared As MyClass`) и функцию доступа `GetShared()`, которая при первом вызове создаёт экземпляр, а потом всегда возвращает один и тот же. Также можно добавить процедуру очистки при выходе.

**Bad:** Класс логгера создаётся в каждой процедуре:

```vba
Public Sub WriteLog(msg As String)
    Dim log As New Logger
    log.Write msg   ' Каждый вызов — новый экземпляр Logger
End Sub
```

**Good:** Один логгер для всего приложения:

```vba
Option Explicit

' Модуль LoggerFactory
Private loggerInstance As Logger

Public Function GetLogger() As Logger
    If loggerInstance Is Nothing Then
        Set loggerInstance = New Logger
    End If
    Set GetLogger = loggerInstance
End Function

Public Sub WriteLog(msg As String)
    Dim log As Logger
    Set log = GetLogger()
    log.Write msg   ' Все используют единый Logger
End Sub
```

> **Примечание:** Это классическая реализация Singleton в VBA. Разместите код в отдельном модуле (как показано), чтобы **любая точка приложения могла получить доступ** к единственному объекту. Не забывайте освобождать `loggerInstance = Nothing` при завершении работы (напр., в основном модуле).

**Summary:** Для единственных по смыслу объектов (например, конфигурация, логгер) применяйте модуль-одиночку: приватная переменная + публичный геттер. Это гарантирует один объект на всё приложение.
**Golden Rules:**

* **Singleton через модуль:** объявите `Private`-переменную и `Public Function GetInstance()`, инициализирующую её при первом обращении.
* **Единый доступ:** все процедуры должны вызывать только `GetShared()` (или `GetInstance`), а не `New Class`.
* **Освобождение:** при закрытии приложения сбрасывайте одиночный объект (`Set objShared = Nothing`).
* **Не злоупотребляйте:** Singleton нужен редко; переоцените, нужны ли глобальные объекты вообще (иногда лучше передавать объекты явно).

**Masterclass:** демонстрирует шаблон Singleton для настраиваемого объекта (например, доступа к базе данных).

```vba
Option Explicit

' Модуль DbManager - обеспечивает единый доступ к БД
Private dbConn As DatabaseConnection

Public Function GetDatabase() As DatabaseConnection
    If dbConn Is Nothing Then
        Set dbConn = New DatabaseConnection
        Call dbConn.Open("Server=...;Database=...")  ' пример настройки
    End If
    Set GetDatabase = dbConn
End Function

Public Sub QueryDatabase(query As String)
    Dim db As DatabaseConnection
    Set db = GetDatabase()
    db.Execute query
End Sub
```

## 4. Структура запуска приложения (Application Bootstrapping)

### 4.1 Единый «Main» модуль приложения

**Проблема:** Начальная последовательность действий (загрузка конфигураций, форм, инициализация бизнес-логики) разрознена. В результате неудобно управлять всего приложением как единым целым.
**Причина:** VBA не требует явного «главного модуля», поэтому разработчики часто пишут инициализационные вызовы в разных местах (событиях, формах).
**Ошибка:** Инициализация рассредоточена: часть кода в `Workbook_Open`, часть – в `UserForm.Initialize`, часть – в не связанных процедурах. Нет одного упорядоченного «Bootstrapping» процесса.
**Решение:** Создайте **центральный модуль** (например, `Module Main`) с одной публичной процедурой `Main()` или `AppStart()`. В ней вызывайте все первичные задачи (загрузка настроек, подключение модулей, вывод главной формы). Затем в `Workbook_Open` просто вызовите `Main`. Это облегчит чтение и поддержку.

**Bad:** Часть инициализации в обработчике открытия, часть в кнопке.

```vba
' В ThisWorkbook:
Private Sub Workbook_Open()
    Call LoadConfig
End Sub

' В другом модуле:
Public Sub LoadConfig()
    ' Загрузка настроек
End Sub

' В UserForm:
Private Sub UserForm_Initialize()
    Call InitializeUI
End Sub
```

**Good:** **Main** модуль с одной процедурой инициализации:

```vba
Option Explicit

' Модуль Main
Public Sub AppStart()
    ' Настроить среду
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' 1) Загрузить настройки
    Call ConfigLoad()

    ' 2) Инициализировать главную форму
    MainForm.Show

    ' 3) Прочие задачи инициализации
    Call SetupEnvironment()

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' ThisWorkbook:
Private Sub Workbook_Open()
    Call AppStart  ' Единый вызов всего процесса
End Sub
```

> **Совет:** Как рекомендуют эксперты, создайте «Main» модуль (не классовый) и соберите все первичные вызовы там. Тогда другие части кода «просто знают», что должны вызывать `AppStart`, инициализация становится упорядоченной.

### 4.2 Настройка окружения (события, вычисления)

**Проблема:** Во время запуска приложения неконтролируемые события (события листов, вычисления) могут сбить инициализацию или ухудшить производительность.
**Причина:** По умолчанию Excel триггерит события при изменениях и персчитывает формулы, пока идёт стартовый код. Если множество макросов вызывают изменение листа, могут сработать лишние обработчики.
**Ошибка:** Программист запускает инициализацию без отключения событий и автоматических пересчётов. Например, в `Workbook_Open` код вызывает макрос, который добавляет лист – это может повторно запустить `Workbook_Open` или `Worksheet_Activate`.
**Решение:** В **начале** инициализации установите `Application.EnableEvents = False` и выставьте оптимальный `Application.Calculation` (например, `xlManual`). После завершения `Application.Calculation = xlAutomatic` и `EnableEvents = True`. Это предотвращает ненужные запуска макросов и ускоряет старт.

**Bad:** Инициализация без контроля:

```vba
Private Sub Workbook_Open()
    ' События всё ещё включены
    LoadData      ' может сработать Worksheet_Activate и т.п.
    Application.Calculation = xlManual
End Sub
```

**Good:** Выключаем и включаем события явно:

```vba
Private Sub Workbook_Open()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    Call LoadData   ' безопасно выполняем загрузку данных

    Application.Calculation = xlAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
```

**Summary:** Чётко определяйте последовательность запуска через `AppStart` или аналог. Помещайте весь код инициализации в единый модуль. Отключайте события и автоматические пересчёты до завершения инициализации.
**Golden Rules:**

* **Main-модуль:** используйте одну публичную `Sub AppStart()` (или `Main`) для всего процесса запуска.
* **Инициализация по порядку:** внутри `AppStart()` последовательно вызывайте загрузку конфигурации, открытие форм, подключение зависимостей и т.д.
* **Контроль событий:** в начале `AppStart` отключите `EnableEvents`, `ScreenUpdating`; включите обратно в конце.
* **Комментирование:** внутри `AppStart()` четко разделяйте шаги инициализации.

**Masterclass:** пример комплексной загрузки приложения: конфигурация → инициализация форм → запуск задач.

```vba
Option Explicit

' Модуль Main:
Public Sub Main()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    ' Загрузка конфигурации
    Call LoadConfig
    
    ' Подключение к базе данных
    Call DatabaseConnect

    ' Инициализация пользовательского интерфейса
    MainForm.Show

    ' Прочие задачи...
    MsgBox "Application initialized."

    Application.Calculation = xlAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub LoadConfig()
    ' Пример загрузки настроек из листа Config
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Config")
    ' ...
End Sub

Private Sub DatabaseConnect()
    ' Пример установки соединения с БД
    Dim conn As Object
    Set conn = New ADODB.Connection
    conn.Open "Provider=SQLOLEDB;Data Source=...;"
    ' Сохраняем соединение в глобальной переменной, если нужно
End Sub
```

## 5. Изоляция глобальных зависимостей

### 5.1 Исключение скрытых «глобалов» (Selection, ActiveWorkbook)

**Проблема:** Код полагается на глобальные объекты и состояние (выделение, активный лист), что усложняет отладку и робастность.
**Причина:** Объекты типа `Selection`, `ActiveCell`, `ActiveWorkbook` меняются с действиями пользователя. Код, который их использует, начинает работать неправильно, если вдруг Excel находится в другом контексте.
**Ошибка:** Процедуры формата обращаются к `Selection` или `ActiveCell`, а процедуры обработки ссылаются на `Application.ActiveWorkbook` без проверки. Это приводит к ошибкам, если окно изменила фокус.
**Решение:** **Изолируйте** доступ к таким объектам. Вместо `Selection` всегда передавайте конкретный диапазон (например, `ByVal rng As Range`). Вместо `ActiveWorkbook` явно используйте нужный `Workbook` (например, `ThisWorkbook`). Если какое-то глобальное состояние неизбежно, оберните его в функцию/сервис.

**Bad:** Применение форматирования к выделению напрямую:

```vba
Public Sub BoldSelection()
    Selection.Font.Bold = True  ' Опасно: что выделено, может быть любым
End Sub
```

**Good:** Форматирование по переданному диапазону:

```vba
Public Sub BoldRange(ByVal rng As Range)
    rng.Font.Bold = True   ' Работает с любым диапазоном, не зависит от выделения
End Sub

' Использование:
BoldRange ThisWorkbook.Worksheets("Data").Range("A1:A10")
```

### 5.2 Избегание глобальных переменных и объектов

**Проблема:** Глобальные переменные (особенно объекты) делают приложение жёстко связанным и трудным для тестирования.
**Причина:** При использовании `Public`-переменных в модулях любая часть кода может изменить состояние, что трудно отследить. Тесты затруднены из-за состояния, сохраняемого между вызовами макросов.
**Ошибка:** Например, в модуле объявлен `Public gWorkbook As Workbook`, инициализируется в одном месте, а используется в другом без проверки – риск нулевой ссылки или неверной книги.
**Решение:** Минимизируйте или уберите глобальные переменные. Если нужен объект для нескольких процедур, внедряйте его параметрами либо передавайте через *Singleton*-геттер (как в секции 3). Для данных используйте локальные или процедуральные переменные.

**Bad:** Глобальная переменная Workbook:

```vba
Public gWB As Workbook

Sub InitWorkbook()
    Set gWB = ThisWorkbook
End Sub

Sub UpdateTitle()
    gWB.Worksheets(1).Name = "Main"  ' Работает только если gWB корректно установлен
End Sub
```

**Good:** Передача книги в процедуру:

```vba
Public Sub UpdateTitle(wb As Workbook)
    wb.Worksheets(1).Name = "Main"
End Sub

' Использование:
Call UpdateTitle(ThisWorkbook)
```

> **Рекомендация:** Идентифицируйте все скрытые глобалы и замените их явными параметрами или обёртками. Например, вместо глобального словаря для настроек используйте функцию доступа (см. раздел 3). Это повышает переносимость кода.

**Summary:** Не допускайте неявной зависимости от глобального состояния Excel. Явно указывайте в коде все объекты (книги, листы, диапазоны). Не используйте `Public`-переменные для данных и объектов, если можно обойтись передачей через параметры или единым сервисом.
**Golden Rules:**

* **Явные параметры:** все процедуры «знают» о своих зависимостях через параметры.
* **Не доверяйте Selection:** замените `Selection`/`ActiveCell` на передаваемые `Range`.
* **Не храните ненужно глобальное состояние:** если что-то нужно во многих процедурах, лучше фабричная функция или Singleton (см. раздел 3).
* **Модульность:** глобальные объекты (Workbook, Worksheet) предпочтительно передавать в функции, а не хранить в `Public`-переменных.

**Masterclass:** пример демонстрирует полную изоляцию глобалов: вместо `Selection` и `ActiveWorkbook` передаём все параметры и используем описанные приёмы.

```vba
Option Explicit

' Централизованная функция обработки данных
Public Sub ProcessData(ByVal dataSheet As Worksheet)
    ' Чётко передаём, какой лист обрабатывать
    Dim rng As Range
    Set rng = dataSheet.Range("B2:B10")
    For Each cell In rng
        cell.Value = cell.Value * 2
    Next cell
End Sub

' Используем в главном коде:
Public Sub RunProcess()
    Dim wb As Workbook
    Set wb = ThisWorkbook  ' или Workbooks("Название.xlsm")
    Call ProcessData(wb.Worksheets("Data"))
End Sub

' Пример безопасности: больше нигде не ссылаемся на ActiveWorkbook/Selection
```

## Итого

В этом разделе мы разобрали ключевые принципы организации старта VBA-приложения: **единая точка входа** (`Workbook_Open`), **централизованная инициализация** (функция `Main`), безопасное использование **единственных объектов (Singleton)** и чёткое **внедрение зависимостей**. На всех этапах необходимо **избегать неявных глобальных ссылок** (ActiveWorkbook, Selection) и распределённой логики. Соблюдение этих правил помогает предотвращать частые архитектурные ошибки, улучшает читаемость и поддержку кода.

## Checklist Rules

* **Must:** Использовать `Workbook_Open` (или аналог) как единственную точку старта приложения.

* **Must:** Всегда объявлять `Option Explicit` и явно указывать типы переменных.

* **Must:** Централизовать инициализацию – собрать её в одном модуле/методе (напр., `Main`).

* **Must:** Выключать события (`EnableEvents = False`) и обновление экрана при запуске, восстанавливать их после инициализации.

* **Must:** Передавать все зависимости (Workbook, Worksheet, объект) через параметры, а не полагаться на `ActiveWorkbook` или `Selection`.

* **Must:** Для уникальных ресурсов (конфигурация, логгер и т.п.) реализовать единственный экземпляр через приватную переменную и геттер (паттерн Singleton).

* **Must Not:** Использовать `Auto_Open` вместо `Workbook_Open`.

* **Must Not:** Хранить важные объекты в `Public`-переменных без крайней необходимости.

* **Must Not:** Разбивать ключевые вызовы инициализации по разным модулям/событиям без порядка.

* **Must Not:** Полагаться на активные книги/листы (`ActiveWorkbook`, `ActiveSheet`) или `Selection` в логике (всегда иметь прямую ссылку на нужный объект).

* **Must Not:** Игнорировать освобождение ресурсов (например, не обнулять объекты при завершении, хотя это в Excel VBA не критично, но желательно).



# Initialization & Entry-point

Первый раздел этого руководства посвящён организации **точки входа** и **инициализации** Excel VBA-приложения. Здесь мы рассмотрим, как задать единый порядок запуска и настроить окружение, а также как централизованно загружать зависимости. Каждый подраздел состоит из: **Проблемы** → **Причины (VBA)** → **Ошибка (пример)** → **Решение (шаблон/приём)**, с наглядными примерами «Bad» → «Good».

## 1. Точка входа и начальная инициализация приложения

### 1.1 Нечёткая точка входа VBA‑макросов

**Проблема:** У Excel нет встроенного явного «Main». Код запускается по событиям (кнопки, события книги) и без централизованного порядка. Это ведёт к разбросу макросов и трудностям управления последовательностью инициализации.
**Причина:** VBA-макросы можно вызывать из разных мест (Auto\_Open, Workbook\_Open, кнопки на листах и т.д.), и без дисциплины их порядок непредсказуем. В частности, использование **Sub Auto\_Open()** считается устаревшим способом запуска макроса при открытии.
**Ошибка:** Код инициализации разбросан по модулям/формам. Например, используют `Auto_Open` или макросы привязанные к кнопкам на разных листах, нет единого порядка. Это приводит к тому, что некоторые процедуры могут не вызвать другие необходимые, а состояние приложения на старте остаётся неопределённым.
**Решение:** Определите одну **точку входа** – например, обработчик события `Workbook_Open` в модуле `ThisWorkbook`, который вызывает процедуру инициализации. Всегда включайте `Option Explicit` для обнаружения ошибок на этапе компиляции. Обновите `Application.EnableEvents`, `ScreenUpdating` и другие свойства вокруг вызова инициализации.

**Bad:** Разрозненный Auto\_Open в стандартном модуле.

```vba
'***** BAD: код запускается через Auto_Open (устарело)
Public Sub Auto_Open()
    ' Установить значение в ячейку текущего активного листа
    ActiveSheet.Range("A1").Value = "Start"
End Sub

Public Sub DoWork()
    ' Пример задачи без гарантии, что инициализация выполнена
    MsgBox "Doing work..."
End Sub
```

**Good:** Единая точка входа – событие `Workbook_Open` в `ThisWorkbook`, вызывающее процедуру инициализации.

```vba
Option Explicit

'***** GOOD: в модуле ThisWorkbook
Private Sub Workbook_Open()
    ' Заблокировать события и анимацию на время инициализации
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Call AppInitialize  ' центральная процедура инициализации

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

'***** GOOD: в стандартном модуле
Public Sub AppInitialize()
    Dim wb As Workbook
    Set wb = ThisWorkbook  ' явно заданная книга для инициализации

    Dim ws As Worksheet
    Set ws = wb.Worksheets("Data")
    ws.Range("A1").Value = "Initialized at " & Now  ' пример инициализации

    ' Дополнительная начальная настройка (напр., режим вычислений):
    Application.Calculation = xlManual  ' отключаем автоматические пересчёты
End Sub
```

> *На будущее:* всегда централизуйте вызов начальных процедур. Используйте `Workbook_Open` для автоматического запуска кода при открытии книги. Всегда `Option Explicit` и обрабатывайте флаги `EnableEvents/ScreenUpdating` вокруг инициализации.

**Summary:** При старте приложения используйте событие `Workbook_Open` (в `ThisWorkbook`) как единую точку входа. Отключайте ненужные события и обновление экрана во время инициализации.
**Golden Rules:**

* **Всегда** включайте `Option Explicit` во всех модулях.
* **Единый запуск:** всю первичную инициализацию вызывайте из `Workbook_Open`.
* **Контроль окружения:** в начале отключайте `ScreenUpdating`, `EnableEvents`; в конце – включайте обратно.
* **Явные объекты:** пользуйтесь `ThisWorkbook`, `Worksheet` и т.д., а не `ActiveSheet`, чтобы быть уверенным, что инициализируется нужная книга/лист.

**Masterclass:** демонстрирует полную последовательную загрузку приложения при открытии книги. Кладёт данные в лист, меняет параметры Excel и восстанавливает состояние.

```vba
Option Explicit

' ThisWorkbook:
Private Sub Workbook_Open()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    MainInitialize  ' главный метод инициализации

    Application.Calculation = xlAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Стандартный модуль:
Public Sub MainInitialize()
    Dim wb As Workbook
    Set wb = ThisWorkbook  ' чётко указываем нужную книгу

    ' Пример: инициализация данных на листах
    Dim wsData As Worksheet
    Set wsData = wb.Worksheets("Data")
    wsData.Range("A1").Value = "App started at " & Now

    Dim wsCfg As Worksheet
    Set wsCfg = wb.Worksheets("Config")
    wsCfg.Range("B1").Value = "Mode: Initialization complete"

    ' Другие действия...
End Sub
```

## 2. Организация загрузки зависимостей и объектов

### 2.1 Неявные глобальные зависимости (ActiveWorkbook/ActiveSheet)

**Проблема:** Процедуры оперируют «скрытыми» объектами VBA (как `ActiveWorkbook`, `ActiveSheet`, `Selection`), что делает код зависимым от текущего состояния Excel.
**Причина:** Если не указать явно, VBA обращается к глобальной среде (например, `Worksheets` равнозначен `Application.ActiveWorkbook.Worksheets`). При этом «Active» может быть любым — пользователь мог переключиться между книгами.
**Ошибка:** Процедуры используют `Worksheets.Add`, `Selection` или `ActiveWorkbook` без явной ссылки. Например, создаётся лист в неожиданной книге или форматируется выделение, не связанное с конкретной книгой.
**Решение:** Всегда передавайте в процедуры нужные объекты (Workbook, Worksheet, Range) параметрами, или используйте явно `ThisWorkbook.Worksheets(…)`. Это называется *внедрением зависимости* (Dependency Injection). Вместо `Worksheets.Add` делайте `wb.Worksheets.Add`, где `wb` — объект книги, переданный в процедуру.

**Bad:** Используется `Worksheets` без ссылки (меняет активную книгу):

```vba
Public Sub CreateReport()
    Dim ws As Worksheet
    Set ws = Worksheets.Add  ' Будет добавлять лист в ActiveWorkbook
    ws.Name = "Report"
End Sub
```

**Good:** Передаём книгу как параметр:

```vba
Public Sub CreateReport(ByVal wb As Workbook)
    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add  ' Явно в заданной книге
    ws.Name = "Report"
End Sub

' В точке вызова:
Call CreateReport(ThisWorkbook)
```

### 2.2 Централизованная инициализация общих объектов

**Проблема:** Один и тот же объект (например, словарь, соединение с БД, фабричный класс) создаётся в разных местах, что дублирует код и расходует ресурсы.
**Причина:** Без выделенного места инициализации каждая процедура «по надобности» делает `New` или `CreateObject`, даже если нужен один экземпляр.
**Ошибка:** Код многократно создаёт объекты. Например, в каждом методе `Dim dict As New Scripting.Dictionary`, приводя к многократному заполнению и разным состояниям.
**Решение:** Создайте **единую точку инициализации** или «фабрику» для общих объектов. Инициализируйте объект один раз и переиспользуйте. Например, заведите `Private dictCache As Object` и функцию `GetDict()`, которая при первом вызове создаст экземпляр, а при последующих — вернёт уже созданный.

**Bad:** В каждой процедуре создаётся новый словарь:

```vba
Public Sub InitData()
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Key", "Value"
    ' ... повторная инициализация каждый раз
End Sub
```

**Good:** Инициализируем один раз и храним в модуле:

```vba
Option Explicit

Private dictData As Object

Public Function GetDataDict() As Object
    If dictData Is Nothing Then
        Set dictData = CreateObject("Scripting.Dictionary")
        dictData.Add "Key", "Value"
    End If
    Set GetDataDict = dictData
End Function

Public Sub UseData()
    Dim d As Object
    Set d = GetDataDict()  ' всегда один и тот же экземпляр
    MsgBox d("Key")
End Sub
```

> **Совет:** Подобный приём – простой вариант шаблона Factory/Singleton. Он устраняет дублирование создания и гарантирует единый объект во всём приложении.

**Summary:** Избегайте «магических» глобальных обращений (`ActiveWorkbook`, `Selection`). Внедряйте необходимые объекты через параметры или централизованные функции. При общих объектах (напр., конфигурация, кэш) создавайте их один раз через приватную переменную в модуле.
**Golden Rules:**

* **Передавайте объекты параметрами** вместо использования `ActiveX`: все процедуры «знают» свои зависимости.
* **Централизуйте создание:** для общих объектов используйте одну функцию или свойство.
* **Избегайте дублирования:** если объект нужен в нескольких местах, храните его единственный экземпляр в модуле.
* **Читаемость:** явно объявляйте типы и используйте Option Explicit, чтобы не полагаться на неявное преобразование.

**Masterclass:** показывает загружаемый синглтон-объект (словарь конфигураций) и использование его в разных частях приложения.

```vba
Option Explicit

' Модуль ConfigManager: обеспечивает доступ к общим настройкам
Private cfg As Object

Public Function GetConfig() As Object
    If cfg Is Nothing Then
        Set cfg = CreateObject("Scripting.Dictionary")
        cfg.Add "Mode", "Test"
        cfg.Add "Version", "1.0"
    End If
    Set GetConfig = cfg
End Function

' Пример использования из разных мест:
Public Sub ShowConfig()
    Dim config As Object
    Set config = GetConfig()
    MsgBox "Mode=" & config("Mode")
End Sub

Public Sub UpdateConfig()
    Dim config As Object
    Set config = GetConfig()
    config("Mode") = "Production"  ' Меняем единственный экземпляр
End Sub
```

## 3. Безопасная реализация Singleton

**Проблема:** Для некоторых классов нужен ровно один экземпляр (конфигуратор, логгер, менеджер ресурсов). VBA не поддерживает приватный конструктор, поэтому легко создать несколько экземпляров случайно.
**Причина:** Класс VBA можно инстанцировать оператором `New` в любом месте. Без контроля можно получить несколько инстансов, что разрушит идею единственного хранилища данных.
**Ошибка:** Объект-одиночка объявляют просто `Public MyObj As New Class`, либо не освобождают. Например, каждый модуль делает `Dim log As New Logger` – получается несколько разных логгеров.
**Решение:** Реализуйте шаблон **Singleton** вручную через модуль. В модуле заведите приватную переменную (`Private objShared As MyClass`) и функцию доступа `GetShared()`, которая при первом вызове создаёт экземпляр, а потом всегда возвращает один и тот же. Также можно добавить процедуру очистки при выходе.

**Bad:** Класс логгера создаётся в каждой процедуре:

```vba
Public Sub WriteLog(msg As String)
    Dim log As New Logger
    log.Write msg   ' Каждый вызов — новый экземпляр Logger
End Sub
```

**Good:** Один логгер для всего приложения:

```vba
Option Explicit

' Модуль LoggerFactory
Private loggerInstance As Logger

Public Function GetLogger() As Logger
    If loggerInstance Is Nothing Then
        Set loggerInstance = New Logger
    End If
    Set GetLogger = loggerInstance
End Function

Public Sub WriteLog(msg As String)
    Dim log As Logger
    Set log = GetLogger()
    log.Write msg   ' Все используют единый Logger
End Sub
```

> **Примечание:** Это классическая реализация Singleton в VBA. Разместите код в отдельном модуле (как показано), чтобы **любая точка приложения могла получить доступ** к единственному объекту. Не забывайте освобождать `loggerInstance = Nothing` при завершении работы (напр., в основном модуле).

**Summary:** Для единственных по смыслу объектов (например, конфигурация, логгер) применяйте модуль-одиночку: приватная переменная + публичный геттер. Это гарантирует один объект на всё приложение.
**Golden Rules:**

* **Singleton через модуль:** объявите `Private`-переменную и `Public Function GetInstance()`, инициализирующую её при первом обращении.
* **Единый доступ:** все процедуры должны вызывать только `GetShared()` (или `GetInstance`), а не `New Class`.
* **Освобождение:** при закрытии приложения сбрасывайте одиночный объект (`Set objShared = Nothing`).
* **Не злоупотребляйте:** Singleton нужен редко; переоцените, нужны ли глобальные объекты вообще (иногда лучше передавать объекты явно).

**Masterclass:** демонстрирует шаблон Singleton для настраиваемого объекта (например, доступа к базе данных).

```vba
Option Explicit

' Модуль DbManager - обеспечивает единый доступ к БД
Private dbConn As DatabaseConnection

Public Function GetDatabase() As DatabaseConnection
    If dbConn Is Nothing Then
        Set dbConn = New DatabaseConnection
        Call dbConn.Open("Server=...;Database=...")  ' пример настройки
    End If
    Set GetDatabase = dbConn
End Function

Public Sub QueryDatabase(query As String)
    Dim db As DatabaseConnection
    Set db = GetDatabase()
    db.Execute query
End Sub
```

## 4. Структура запуска приложения (Application Bootstrapping)

### 4.1 Единый «Main» модуль приложения

**Проблема:** Начальная последовательность действий (загрузка конфигураций, форм, инициализация бизнес-логики) разрознена. В результате неудобно управлять всего приложением как единым целым.
**Причина:** VBA не требует явного «главного модуля», поэтому разработчики часто пишут инициализационные вызовы в разных местах (событиях, формах).
**Ошибка:** Инициализация рассредоточена: часть кода в `Workbook_Open`, часть – в `UserForm.Initialize`, часть – в не связанных процедурах. Нет одного упорядоченного «Bootstrapping» процесса.
**Решение:** Создайте **центральный модуль** (например, `Module Main`) с одной публичной процедурой `Main()` или `AppStart()`. В ней вызывайте все первичные задачи (загрузка настроек, подключение модулей, вывод главной формы). Затем в `Workbook_Open` просто вызовите `Main`. Это облегчит чтение и поддержку.

**Bad:** Часть инициализации в обработчике открытия, часть в кнопке.

```vba
' В ThisWorkbook:
Private Sub Workbook_Open()
    Call LoadConfig
End Sub

' В другом модуле:
Public Sub LoadConfig()
    ' Загрузка настроек
End Sub

' В UserForm:
Private Sub UserForm_Initialize()
    Call InitializeUI
End Sub
```

**Good:** **Main** модуль с одной процедурой инициализации:

```vba
Option Explicit

' Модуль Main
Public Sub AppStart()
    ' Настроить среду
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' 1) Загрузить настройки
    Call ConfigLoad()

    ' 2) Инициализировать главную форму
    MainForm.Show

    ' 3) Прочие задачи инициализации
    Call SetupEnvironment()

    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' ThisWorkbook:
Private Sub Workbook_Open()
    Call AppStart  ' Единый вызов всего процесса
End Sub
```

> **Совет:** Как рекомендуют эксперты, создайте «Main» модуль (не классовый) и соберите все первичные вызовы там. Тогда другие части кода «просто знают», что должны вызывать `AppStart`, инициализация становится упорядоченной.

### 4.2 Настройка окружения (события, вычисления)

**Проблема:** Во время запуска приложения неконтролируемые события (события листов, вычисления) могут сбить инициализацию или ухудшить производительность.
**Причина:** По умолчанию Excel триггерит события при изменениях и персчитывает формулы, пока идёт стартовый код. Если множество макросов вызывают изменение листа, могут сработать лишние обработчики.
**Ошибка:** Программист запускает инициализацию без отключения событий и автоматических пересчётов. Например, в `Workbook_Open` код вызывает макрос, который добавляет лист – это может повторно запустить `Workbook_Open` или `Worksheet_Activate`.
**Решение:** В **начале** инициализации установите `Application.EnableEvents = False` и выставьте оптимальный `Application.Calculation` (например, `xlManual`). После завершения `Application.Calculation = xlAutomatic` и `EnableEvents = True`. Это предотвращает ненужные запуска макросов и ускоряет старт.

**Bad:** Инициализация без контроля:

```vba
Private Sub Workbook_Open()
    ' События всё ещё включены
    LoadData      ' может сработать Worksheet_Activate и т.п.
    Application.Calculation = xlManual
End Sub
```

**Good:** Выключаем и включаем события явно:

```vba
Private Sub Workbook_Open()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    Call LoadData   ' безопасно выполняем загрузку данных

    Application.Calculation = xlAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
```

**Summary:** Чётко определяйте последовательность запуска через `AppStart` или аналог. Помещайте весь код инициализации в единый модуль. Отключайте события и автоматические пересчёты до завершения инициализации.
**Golden Rules:**

* **Main-модуль:** используйте одну публичную `Sub AppStart()` (или `Main`) для всего процесса запуска.
* **Инициализация по порядку:** внутри `AppStart()` последовательно вызывайте загрузку конфигурации, открытие форм, подключение зависимостей и т.д.
* **Контроль событий:** в начале `AppStart` отключите `EnableEvents`, `ScreenUpdating`; включите обратно в конце.
* **Комментирование:** внутри `AppStart()` четко разделяйте шаги инициализации.

**Masterclass:** пример комплексной загрузки приложения: конфигурация → инициализация форм → запуск задач.

```vba
Option Explicit

' Модуль Main:
Public Sub Main()
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    ' Загрузка конфигурации
    Call LoadConfig
    
    ' Подключение к базе данных
    Call DatabaseConnect

    ' Инициализация пользовательского интерфейса
    MainForm.Show

    ' Прочие задачи...
    MsgBox "Application initialized."

    Application.Calculation = xlAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub LoadConfig()
    ' Пример загрузки настроек из листа Config
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Config")
    ' ...
End Sub

Private Sub DatabaseConnect()
    ' Пример установки соединения с БД
    Dim conn As Object
    Set conn = New ADODB.Connection
    conn.Open "Provider=SQLOLEDB;Data Source=...;"
    ' Сохраняем соединение в глобальной переменной, если нужно
End Sub
```

## 5. Изоляция глобальных зависимостей

### 5.1 Исключение скрытых «глобалов» (Selection, ActiveWorkbook)

**Проблема:** Код полагается на глобальные объекты и состояние (выделение, активный лист), что усложняет отладку и робастность.
**Причина:** Объекты типа `Selection`, `ActiveCell`, `ActiveWorkbook` меняются с действиями пользователя. Код, который их использует, начинает работать неправильно, если вдруг Excel находится в другом контексте.
**Ошибка:** Процедуры формата обращаются к `Selection` или `ActiveCell`, а процедуры обработки ссылаются на `Application.ActiveWorkbook` без проверки. Это приводит к ошибкам, если окно изменила фокус.
**Решение:** **Изолируйте** доступ к таким объектам. Вместо `Selection` всегда передавайте конкретный диапазон (например, `ByVal rng As Range`). Вместо `ActiveWorkbook` явно используйте нужный `Workbook` (например, `ThisWorkbook`). Если какое-то глобальное состояние неизбежно, оберните его в функцию/сервис.

**Bad:** Применение форматирования к выделению напрямую:

```vba
Public Sub BoldSelection()
    Selection.Font.Bold = True  ' Опасно: что выделено, может быть любым
End Sub
```

**Good:** Форматирование по переданному диапазону:

```vba
Public Sub BoldRange(ByVal rng As Range)
    rng.Font.Bold = True   ' Работает с любым диапазоном, не зависит от выделения
End Sub

' Использование:
BoldRange ThisWorkbook.Worksheets("Data").Range("A1:A10")
```

### 5.2 Избегание глобальных переменных и объектов

**Проблема:** Глобальные переменные (особенно объекты) делают приложение жёстко связанным и трудным для тестирования.
**Причина:** При использовании `Public`-переменных в модулях любая часть кода может изменить состояние, что трудно отследить. Тесты затруднены из-за состояния, сохраняемого между вызовами макросов.
**Ошибка:** Например, в модуле объявлен `Public gWorkbook As Workbook`, инициализируется в одном месте, а используется в другом без проверки – риск нулевой ссылки или неверной книги.
**Решение:** Минимизируйте или уберите глобальные переменные. Если нужен объект для нескольких процедур, внедряйте его параметрами либо передавайте через *Singleton*-геттер (как в секции 3). Для данных используйте локальные или процедуральные переменные.

**Bad:** Глобальная переменная Workbook:

```vba
Public gWB As Workbook

Sub InitWorkbook()
    Set gWB = ThisWorkbook
End Sub

Sub UpdateTitle()
    gWB.Worksheets(1).Name = "Main"  ' Работает только если gWB корректно установлен
End Sub
```

**Good:** Передача книги в процедуру:

```vba
Public Sub UpdateTitle(wb As Workbook)
    wb.Worksheets(1).Name = "Main"
End Sub

' Использование:
Call UpdateTitle(ThisWorkbook)
```

> **Рекомендация:** Идентифицируйте все скрытые глобалы и замените их явными параметрами или обёртками. Например, вместо глобального словаря для настроек используйте функцию доступа (см. раздел 3). Это повышает переносимость кода.

**Summary:** Не допускайте неявной зависимости от глобального состояния Excel. Явно указывайте в коде все объекты (книги, листы, диапазоны). Не используйте `Public`-переменные для данных и объектов, если можно обойтись передачей через параметры или единым сервисом.
**Golden Rules:**

* **Явные параметры:** все процедуры «знают» о своих зависимостях через параметры.
* **Не доверяйте Selection:** замените `Selection`/`ActiveCell` на передаваемые `Range`.
* **Не храните ненужно глобальное состояние:** если что-то нужно во многих процедурах, лучше фабричная функция или Singleton (см. раздел 3).
* **Модульность:** глобальные объекты (Workbook, Worksheet) предпочтительно передавать в функции, а не хранить в `Public`-переменных.

**Masterclass:** пример демонстрирует полную изоляцию глобалов: вместо `Selection` и `ActiveWorkbook` передаём все параметры и используем описанные приёмы.

```vba
Option Explicit

' Централизованная функция обработки данных
Public Sub ProcessData(ByVal dataSheet As Worksheet)
    ' Чётко передаём, какой лист обрабатывать
    Dim rng As Range
    Set rng = dataSheet.Range("B2:B10")
    For Each cell In rng
        cell.Value = cell.Value * 2
    Next cell
End Sub

' Используем в главном коде:
Public Sub RunProcess()
    Dim wb As Workbook
    Set wb = ThisWorkbook  ' или Workbooks("Название.xlsm")
    Call ProcessData(wb.Worksheets("Data"))
End Sub

' Пример безопасности: больше нигде не ссылаемся на ActiveWorkbook/Selection
```

## Итого

В этом разделе мы разобрали ключевые принципы организации старта VBA-приложения: **единая точка входа** (`Workbook_Open`), **централизованная инициализация** (функция `Main`), безопасное использование **единственных объектов (Singleton)** и чёткое **внедрение зависимостей**. На всех этапах необходимо **избегать неявных глобальных ссылок** (ActiveWorkbook, Selection) и распределённой логики. Соблюдение этих правил помогает предотвращать частые архитектурные ошибки, улучшает читаемость и поддержку кода.

## Checklist Rules

* **Must:** Использовать `Workbook_Open` (или аналог) как единственную точку старта приложения.

* **Must:** Всегда объявлять `Option Explicit` и явно указывать типы переменных.

* **Must:** Централизовать инициализацию – собрать её в одном модуле/методе (напр., `Main`).

* **Must:** Выключать события (`EnableEvents = False`) и обновление экрана при запуске, восстанавливать их после инициализации.

* **Must:** Передавать все зависимости (Workbook, Worksheet, объект) через параметры, а не полагаться на `ActiveWorkbook` или `Selection`.

* **Must:** Для уникальных ресурсов (конфигурация, логгер и т.п.) реализовать единственный экземпляр через приватную переменную и геттер (паттерн Singleton).

* **Must Not:** Использовать `Auto_Open` вместо `Workbook_Open`.

* **Must Not:** Хранить важные объекты в `Public`-переменных без крайней необходимости.

* **Must Not:** Разбивать ключевые вызовы инициализации по разным модулям/событиям без порядка.

* **Must Not:** Полагаться на активные книги/листы (`ActiveWorkbook`, `ActiveSheet`) или `Selection` в логике (всегда иметь прямую ссылку на нужный объект).

* **Must Not:** Игнорировать освобождение ресурсов (например, не обнулять объекты при завершении, хотя это в Excel VBA не критично, но желательно).
