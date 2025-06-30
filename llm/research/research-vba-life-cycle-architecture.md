# Проектирование архитектуры приложений VBA: Инициализация и Точка входа

## Содержание

1.  [Введение](#введение)
    *   [Цель руководства](#цель-руководства)
    *   [Важность правильной инициализации в VBA](#важность-правильной-инициализации-в-vba)
    *   [Обзор типичных проблем ИИ-ассистентов](#обзор-типичных-проблем-ии-ассистентов)
2.  [Раздел 1: Основы инициализации в VBA](#раздел-1-основы-инициализации-в-vba)
    *   [1.1. Подраздел: Понимание контекста выполнения VBA](#11-подраздел-понимание-контекста-выполнения-vba)
    *   [1.2. Подраздел: Точки входа приложения VBA](#12-подраздел-точки-входа-приложения-vba)
    *   [Резюме Раздела 1](#резюме-раздела-1)
    *   [Золотые правила Раздела 1](#золотые-правила-раздела-1)
    *   [Мастер-класс Раздела 1](#мастер-класс-раздела-1)
3.  [Раздел 2: Стратегии и шаблоны инициализации](#раздел-2-стратегии-и-шаблоны-инициализации)
    *   [2.1. Подраздел: Bootstrapping (Начальная загрузка) приложения](#21-подраздел-bootstrapping-начальная-загрузка-приложения)
    *   [2.2. Подраздел: Инициализация глобальных и разделяемых ресурсов (Singleton)](#22-подраздел-инициализация-глобальных-и-разделяемых-ресурсов-singleton)
    *   [2.3. Подраздел: Отложенная (Lazy) инициализация](#23-подраздел-отложенная-lazy-инициализация)
    *   [Резюме Раздела 2](#резюме-раздела-2)
    *   [Золотые правила Раздела 2](#золотые-правила-раздела-2)
    *   [Мастер-класс Раздела 2](#мастер-класс-раздела-2)
4.  [Раздел 3: Организация доступа к функциональности](#раздел-3-организация-доступа-к-функциональности)
    *   [3.1. Подраздел: Использование стандартных модулей vs. модулей классов](#31-подраздел-использование-стандартных-модулей-vs-модулей-классов)
    *   [3.2. Подраздел: Фасады для упрощения взаимодействия](#32-подраздел-фасады-для-упрощения-взаимодействия)
    *   [Резюме Раздела 3](#резюме-раздела-3)
    *   [Золотые правила Раздела 3](#золотые-правила-раздела-3)
    *   [Мастер-класс Раздела 3](#мастер-класс-раздела-3)
5.  [Заключение](#заключение)
6.  [Список использованных источников (иллюстративный)](#список-использованных-источников-иллюстративный)

---

## Введение

### Цель руководства
Данное руководство предназначено для ИИ-ассистентов и разработчиков VBA с целью минимизации архитектурных ошибок при генерации и написании кода VBA. Основное внимание уделяется правильным подходам к инициализации компонентов приложения, управлению точками входа и организации доступа к функциональности, учитывая специфику и ограничения VBA.

### Важность правильной инициализации в VBA
VBA, будучи событийно-ориентированным и интерпретируемым языком, встроенным в хост-приложения (например, Excel, Word), требует особого подхода к инициализации. Ошибки на этом этапе приводят к нестабильной работе, трудностям в отладке, утечкам памяти и проблемам с управлением состоянием приложения. Правильная архитектура инициализации – фундамент для надежных и масштабируемых VBA-решений.

### Обзор типичных проблем ИИ-ассистентов
ИИ-ассистенты часто допускают следующие ошибки из-за недостаточного понимания контекста VBA:
*   Применение паттернов из VB.NET или других ООП-языков без адаптации к VBA.
*   Некорректная реализация глобального доступа к объектам (например, Singleton).
*   Ошибки в определении и управлении точками входа приложения.
*   Проблемы с порядком инициализации зависимых компонентов.
*   Неправильное использование стандартных модулей и модулей классов для инициализации.

---

## Раздел 1: Основы инициализации в VBA

### 1.1. Подраздел: Понимание контекста выполнения VBA

*   **Проблема:** Неверное представление о запуске и жизненном цикле VBA-кода, особенно в отношении автоматической инициализации объектов.
*   **Причина (в контексте VBA):** VBA-код обычно запускается событием (например, `Workbook_Open`, нажатие кнопки на листе) или прямым вызовом макроса пользователем. В VBA отсутствует концепция единого метода `Main()`, который автоматически выполняется при запуске приложения, как в некоторых других языках программирования. Переменные уровня модуля (включая объектные) в стандартных модулях сбрасываются при возникновении необработанной ошибки или при явном сбросе проекта (End, Debug -> Stop, изменение кода в режиме выполнения). В модулях классов переменные существуют, пока существует экземпляр класса.
*   **Ошибка (специфический симптом/пример):** Попытка объявить `Public` объектную переменную в стандартном модуле и ожидать, что она будет автоматически создана (`New`) и доступна сразу после открытия книги без явной инициализации. Это приводит к ошибке "Object variable or With block variable not set" (Ошибка 91) при первом обращении.
    ````vba
    ' В стандартном модуле (Module1)
    Public g_AppHandler As Object ' Ожидается, что это будет экземпляр класса CAppHandler
    
    Sub TestGlobalHandler()
        ' Ошибка 91 здесь, если g_AppHandler не был инициализирован
        g_AppHandler.DoSomething 
    End Sub
    ````
*   **Решение (VBA-специфический паттерн/подход):** Использовать явные процедуры инициализации, которые создают экземпляры необходимых объектов. Эти процедуры должны вызываться из четко определенных точек входа (например, `Workbook_Open`).
*   **Примеры кода:**

    **Плохо (Bad):**
    ````vba
    ' Module1 - Стандартный модуль
    Public g_MySheetHelper As CSheetHelper ' Объявлено, но не инициализировано
    
    Sub UseHelper()
        ' Это вызовет ошибку 91, если InitializeHelper не был вызван ранее
        Debug.Print g_MySheetHelper.GetSheetName(ThisWorkbook.Worksheets(1))
    End Sub
    
    ' CSheetHelper - Модуль класса
    ' Option Explicit
    ' Public Function GetSheetName(sht As Worksheet) As String
    '     GetSheetName = sht.Name
    ' End Function
    ````

    **Хорошо (Good):**
    ````vba
    ' Module1 - Стандартный модуль
    Public g_MySheetHelper As CSheetHelper
    
    Sub InitializeHelper()
        Set g_MySheetHelper = New CSheetHelper
        Debug.Print "CSheetHelper инициализирован."
    End Sub
    
    Sub UseHelper()
        If g_MySheetHelper Is Nothing Then
            Debug.Print "Helper не инициализирован. Вызов InitializeHelper..."
            InitializeHelper
        End If
        Debug.Print g_MySheetHelper.GetSheetName(ThisWorkbook.Worksheets(1))
    End Sub
    
    ' Вызывать InitializeHelper из точки входа, например, Workbook_Open
    ' В ThisWorkbook:
    ' Private Sub Workbook_Open()
    '     Call Module1.InitializeHelper
    ' End Sub

    ' CSheetHelper - Модуль класса
    ' Option Explicit
    ' Public Function GetSheetName(sht As Worksheet) As String
    '     GetSheetName = sht.Name
    ' End Function
    ````
*   **Резюме:** Объекты в VBA требуют явной инициализации (`Set New ...`). Не полагайтесь на автоматическое создание экземпляров при объявлении, особенно для глобально доступных объектов.
*   **Контрольные правила:**
    *   `Must`: Явно инициализировать все объектные переменные перед использованием.
    *   `Must Not`: Ожидать автоматической инициализации объектов, объявленных в стандартных модулях, при загрузке проекта.

### 1.2. Подраздел: Точки входа приложения VBA

*   **Проблема:** Неопределенность в выборе и управлении начальной точкой (или точками) выполнения основной логики приложения, что затрудняет контроль над процессом инициализации.
*   **Причина (в контексте VBA):** VBA предоставляет множество потенциальных точек входа:
    *   События книги: `Workbook_Open`, `Workbook_Activate`.
    *   События листа: `Worksheet_Activate`, `Worksheet_Change`.
    *   События элементов управления: `CommandButton_Click`.
    *   Автоматически выполняемые процедуры: `Auto_Open`, `Auto_Close` (менее предпочтительны, чем события книги).
    *   Запуск макроса пользователем (Alt+F8).
    *   События пользовательских форм: `UserForm_Initialize`.
*   **Ошибка (специфический симптом/пример):** Размещение критически важной логики инициализации в нескольких разных обработчиках событий (например, часть в `Workbook_Open`, часть в `Worksheet_Activate` первого листа, часть при нажатии кнопки). Это приводит к:
    *   Повторной или неполной инициализации.
    *   Трудностям в отслеживании состояния приложения.
    *   Непредсказуемому поведению, если пользователь взаимодействует с книгой не так, как ожидал разработчик (например, открывает другой лист первым).
*   **Решение (VBA-специфический паттерн/подход):**
    1.  Определить основную, единую процедуру-инициализатор (часто называемую `InitializeApplication`, `Startup` или `Bootstrap`).
    2.  Вызывать эту процедуру из главной предполагаемой точки входа, чаще всего `Workbook_Open`.
    3.  Обеспечить, чтобы инициализация выполнялась только один раз, если это необходимо (например, с помощью статической переменной-флага или проверки состояния).
*   **Примеры кода:**

    **Плохо (Bad):**
    ````vba
    ' В ThisWorkbook
    Private Sub Workbook_Open()
        ' Инициализация настроек
        Call LoadSettings
        Debug.Print "Настройки загружены из Workbook_Open"
    End Sub

    ' В Module1
    Public Sub LoadSettings()
        ' ... логика загрузки настроек ...
    End Sub

    Public Sub ProcessData()
        ' Предполагается, что настройки уже загружены
        If IsSettingsLoaded = False Then ' IsSettingsLoaded - некая глобальная переменная
            Debug.Print "Настройки не загружены при вызове ProcessData!"
            ' Возможно, попытка загрузить их здесь, что нарушает централизацию
        End If
        ' ... работа с данными ...
    End Sub
    ````

    **Хорошо (Good):**
    ````vba
    ' В ThisWorkbook
    Private Sub Workbook_Open()
        ApplicationController.Startup ThisWorkbook
    End Sub

    ' Module: ApplicationController (Стандартный модуль)
    Option Explicit
    Private m_IsInitialized As Boolean
    Public g_AppSettings As Object ' Например, Scripting.Dictionary

    Public Sub Startup(ByVal wb As Workbook)
        If m_IsInitialized Then
            Debug.Print "Приложение уже инициализировано."
            Exit Sub
        End If
        
        Debug.Print "Запуск инициализации приложения..."
        
        ' 1. Загрузка настроек
        Set g_AppSettings = LoadAppSettings(wb)
        If g_AppSettings Is Nothing Then
            MsgBox "Не удалось загрузить настройки приложения. Работа будет прекращена.", vbCritical
            ' Здесь можно добавить логику закрытия книги или другие действия
            Exit Sub
        End If
        Debug.Print "Настройки приложения загружены."
        
        ' 2. Инициализация других сервисов (логгер, менеджер данных и т.д.)
        ' Call InitializeLogger
        ' Call InitializeDataManager
        
        m_IsInitialized = True
        Debug.Print "Приложение успешно инициализировано."
        MsgBox "Добро пожаловать! Приложение готово к работе.", vbInformation
    End Sub

    Private Function LoadAppSettings(ByVal wb As Workbook) As Object ' Scripting.Dictionary
        ' Пример: загрузка настроек с отдельного листа "Settings"
        Dim wsSettings As Worksheet
        Dim settingsDict As Object ' Scripting.Dictionary
        Set settingsDict = CreateObject("Scripting.Dictionary")
        
        On Error Resume Next
        Set wsSettings = wb.Worksheets("Settings")
        On Error GoTo 0
        
        If wsSettings Is Nothing Then
            Debug.Print "Лист 'Settings' не найден."
            Set LoadAppSettings = Nothing
            Exit Function
        End If
        
        ' Пример чтения: ключ в колонке A, значение в колонке B
        Dim lRow As Long
        For lRow = 1 To wsSettings.Cells(Rows.Count, "A").End(xlUp).Row
            If Trim(wsSettings.Cells(lRow, "A").Value) <> "" Then
                settingsDict(Trim(wsSettings.Cells(lRow, "A").Value)) = wsSettings.Cells(lRow, "B").Value
            End If
        Next lRow
        
        Set LoadAppSettings = settingsDict
    End Function

    Public Sub Shutdown()
        If Not m_IsInitialized Then Exit Sub
        
        Debug.Print "Деинициализация приложения..."
        ' Освобождение ресурсов
        Set g_AppSettings = Nothing
        ' Call DeinitializeLogger
        ' Call DeinitializeDataManager
        
        m_IsInitialized = False
        Debug.Print "Приложение деинициализировано."
    End Sub
    
    ' В ThisWorkbook также можно добавить:
    ' Private Sub Workbook_BeforeClose(Cancel As Boolean)
    '     ApplicationController.Shutdown
    ' End Sub
    ````
*   **Резюме:** Централизуйте логику инициализации приложения в одной процедуре, вызываемой из основной точки входа (обычно `Workbook_Open`). Это улучшает управляемость, предсказуемость и облегчает отладку.
*   **Контрольные правила:**
    *   `Must`: Определить единую процедуру для инициализации приложения.
    *   `Must`: Вызывать процедуру инициализации из основной точки входа (например, `Workbook_Open`).
    *   `Must Not`: Разбрасывать логику инициализации по нескольким несвязанным событиям или процедурам.
    *   `Must`: Рассмотреть возможность добавления процедуры деинициализации (`Shutdown`), вызываемой, например, из `Workbook_BeforeClose`.

### Резюме Раздела 1
Правильная инициализация в VBA начинается с понимания его событийно-управляемой природы и отсутствия стандартного `Main()`. Ключевым является явная инициализация всех объектов и централизация логики запуска приложения через единую точку входа, обычно `Workbook_Open`, которая вызывает специализированную процедуру-инициализатор.

### Золотые правила Раздела 1
1.  **Явность превыше всего:** Всегда явно создавайте экземпляры объектов (`Set obj = New ClassName`).
2.  **Единая точка входа:** Используйте `Workbook_Open` (или аналогичную основную точку) для запуска централизованной процедуры инициализации.
3.  **Контроль инициализации:** Убедитесь, что критическая инициализация происходит один раз и в правильном порядке.
4.  **Готовность к ошибкам:** Предусматривайте возможные сбои при инициализации (например, отсутствие конфигурационных файлов или листов) и обрабатывайте их корректно.

### Мастер-класс Раздела 1: Простое приложение с четкой точкой входа и инициализацией

**Сценарий:** Приложение Excel, которое при открытии проверяет наличие определенного листа ("DataSheet") и, если он существует, инициализирует простой "Менеджер Данных", который отображает количество записей на этом листе.

**1. Модуль класса `CDataManager`**
````vba
' CDataManager - Модуль класса
Option Explicit

Private m_DataSheet As Worksheet
Private m_RecordCount As Long

Public Sub Initialize(ByVal targetSheet As Worksheet)
    If targetSheet Is Nothing Then
        Debug.Print "Ошибка инициализации CDataManager: лист не указан."
        Exit Sub
    End If
    Set m_DataSheet = targetSheet
    RefreshRecordCount
    Debug.Print "CDataManager инициализирован для листа: " & m_DataSheet.Name
End Sub

Public Sub RefreshRecordCount()
    If m_DataSheet Is Nothing Then Exit Sub
    ' Простое определение количества строк с данными в столбце A
    m_RecordCount = m_DataSheet.Cells(Rows.Count, "A").End(xlUp).Row
    ' Если лист пуст, но есть заголовок, End(xlUp) может вернуть 1.
    ' Уточним, если первая ячейка пуста, считаем 0 записей.
    If m_RecordCount = 1 And IsEmpty(m_DataSheet.Cells(1, "A").Value) Then
        m_RecordCount = 0
    End If
End Sub

Public Function GetRecordCount() As Long
    GetRecordCount = m_RecordCount
End Function

Public Property Get DataSheetName() As String
    If Not m_DataSheet Is Nothing Then
        DataSheetName = m_DataSheet.Name
    Else
        DataSheetName = "N/A"
    End If
End Property

Private Sub Class_Terminate()
    Set m_DataSheet = Nothing
    Debug.Print "CDataManager деинициализирован."
End Sub
````

**2. Стандартный модуль `AppMain`**
````vba
' AppMain - Стандартный модуль
Option Explicit

Public g_DataManager As CDataManager

Sub InitializeApplication(ByVal wb As Workbook)
    Dim dataSheet As Worksheet
    
    On Error Resume Next
    Set dataSheet = wb.Worksheets("DataSheet")
    On Error GoTo 0
    
    If dataSheet Is Nothing Then
        MsgBox "Лист 'DataSheet' не найден. Функциональность обработки данных будет недоступна.", vbExclamation
        Set g_DataManager = Nothing ' Явно указываем, что менеджер не создан
    Else
        Set g_DataManager = New CDataManager
        g_DataManager.Initialize dataSheet
        MsgBox "Менеджер данных инициализирован. Найдено записей: " & g_DataManager.GetRecordCount(), vbInformation
    End If
End Sub

Sub ShowRecordCount()
    If g_DataManager Is Nothing Then
        MsgBox "Менеджер данных не инициализирован. Возможно, лист 'DataSheet' отсутствует.", vbCritical
        Exit Sub
    End If
    
    g_DataManager.RefreshRecordCount ' Обновить на случай изменений
    MsgBox "Текущее количество записей на листе '" & g_DataManager.DataSheetName & "': " & g_DataManager.GetRecordCount(), vbInformation
End Sub

Sub DeinitializeApplication()
    Set g_DataManager = Nothing
    Debug.Print "Приложение деинициализировано в AppMain."
End Sub
````

**3. Модуль `ThisWorkbook`**
````vba
' ThisWorkbook - Модуль книги
Option Explicit

Private Sub Workbook_Open()
    Call AppMain.InitializeApplication(Me)
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call AppMain.DeinitializeApplication
    Debug.Print "Книга закрывается."
End Sub
````

**Как это работает:**
1.  При открытии книги Excel срабатывает событие `Workbook_Open`.
2.  `Workbook_Open` вызывает `AppMain.InitializeApplication`, передавая текущий объект `Workbook`.
3.  `InitializeApplication` пытается найти лист "DataSheet".
4.  Если лист найден, создается экземпляр `CDataManager`, который инициализируется этим листом и подсчитывает записи. Глобальная переменная `g_DataManager` теперь содержит ссылку на этот объект.
5.  Если лист не найден, выводится сообщение, и `g_DataManager` остается `Nothing`.
6.  Пользователь может вызвать макрос `ShowRecordCount` (например, с кнопки), который использует `g_DataManager`.
7.  При закрытии книги `Workbook_BeforeClose` вызывает `AppMain.DeinitializeApplication`, которая освобождает объект `g_DataManager` (`Set = Nothing`), что, в свою очередь, вызывает `Class_Terminate` для `CDataManager`.

Этот пример демонстрирует четкую точку входа, централизованную инициализацию и базовое управление состоянием.

---

## Раздел 2: Стратегии и шаблоны инициализации

### 2.1. Подраздел: Bootstrapping (Начальная загрузка) приложения

*   **Проблема:** Отсутствие централизованного и упорядоченного механизма для настройки и запуска всех основных компонентов и сервисов приложения при его старте.
*   **Причина (в контексте VBA):** По мере роста сложности приложения количество объектов, которые нужно инициализировать (логгеры, конфигураторы, обработчики данных, UI-менеджеры), увеличивается. Без специального координатора процесс инициализации становится хаотичным и трудноуправляемым.
*   **Ошибка (специфический симптом/пример):**
    *   Логика инициализации размазана по нескольким модулям или процедурам.
    *   Нарушение порядка инициализации зависимостей (например, попытка использовать логгер до его инициализации).
    *   Трудности при добавлении новых сервисов, так как приходится изменять множество мест.
    *   Процедура `Workbook_Open` становится слишком большой и сложной.
*   **Решение (VBA-специфический паттерн/подход):** Реализация модуля или класса "Bootstrapper" (Загрузчик). Этот компонент отвечает за весь процесс начальной загрузки:
    1.  Определение последовательности инициализации.
    2.  Создание и настройка экземпляров всех ключевых сервисов.
    3.  Обработка ошибок на этапе инициализации.
    Bootstrapper обычно вызывается из основной точки входа (например, `Workbook_Open`).
*   **Примеры кода:**

    **Плохо (Bad):** `Workbook_Open` делает слишком много.
    ````vba
    ' В ThisWorkbook
    Private Sub Workbook_Open()
        ' 1. Инициализация логгера
        Call InitializeMyLogger(Me.Path & "\app.log")
        LogMessage "Приложение запущено."
        
        ' 2. Загрузка конфигурации
        Dim configPath As String
        configPath = Me.Path & "\config.ini"
        If Not FileExists(configPath) Then
            LogMessage "Файл конфигурации не найден: " & configPath, vbCritical
            MsgBox "Ошибка: Файл конфигурации отсутствует!", vbCritical
            Exit Sub ' Проблема: частичная инициализация, логгер работает
        End If
        Call LoadConfigurationFromFile(configPath)
        LogMessage "Конфигурация загружена."
        
        ' 3. Инициализация менеджера данных
        Call InitializeDataManager(GetConfigValue("DBPath"))
        LogMessage "Менеджер данных инициализирован."
        
        ' ... и так далее ...
    End Sub
    
    ' Функции InitializeMyLogger, LogMessage, FileExists, LoadConfigurationFromFile, GetConfigValue, InitializeDataManager
    ' разбросаны по разным модулям или находятся здесь же, делая код громоздким.
    ````

    **Хорошо (Good):** Использование модуля `AppBootstrapper`.
    ````vba
    ' AppBootstrapper - Стандартный модуль
    Option Explicit

    Private m_IsBootstrapped As Boolean
    Public g_Logger As Object ' CLogger
    Public g_Config As Object ' CConfig
    Public g_DataManager As Object ' CDataManager

    Public Sub Execute(ByVal wb As Workbook)
        If m_IsBootstrapped Then
            Debug.Print "Bootstrapper: Уже выполнен."
            Exit Sub
        End If

        Debug.Print "Bootstrapper: Запуск..."

        ' Шаг 1: Инициализация Логгера (минимальные зависимости)
        On Error GoTo BootstrapErrorHandler
        Set g_Logger = New CLogger
        g_Logger.Initialize wb.Path & "\application.log"
        g_Logger.Info "Bootstrapper: Логгер инициализирован."

        ' Шаг 2: Загрузка Конфигурации
        Set g_Config = New CConfig
        g_Config.LoadFromFile wb.Path & "\settings.xml"
        g_Logger.Info "Bootstrapper: Конфигурация загружена."

        ' Шаг 3: Инициализация Менеджера Данных (зависит от Конфигурации)
        Set g_DataManager = New CDataManager
        g_DataManager.Initialize g_Config.GetSetting("DatabaseConnectionString")
        g_Logger.Info "Bootstrapper: Менеджер данных инициализирован."
        
        ' Шаг 4: Инициализация UI или других компонентов
        ' Call InitializeUserInterface

        m_IsBootstrapped = True
        g_Logger.Info "Bootstrapper: Успешно завершен."
        MsgBox "Приложение успешно инициализировано и готово к работе!", vbInformation
        Exit Sub

    BootstrapErrorHandler:
        Dim errorMsg As String
        errorMsg = "Критическая ошибка при инициализации приложения: " & Err.Description & _
                   " (Источник: " & Err.Source & ", Номер: " & Err.Number & ")"
        Debug.Print errorMsg
        If Not g_Logger Is Nothing Then
            g_Logger.Error errorMsg
        End If
        MsgBox errorMsg, vbCritical, "Ошибка инициализации"
        ' Здесь можно добавить логику аварийного завершения или очистки
        Call CleanUpPartialBootstrap
        m_IsBootstrapped = False ' Сброс флага
    End Sub

    Public Sub CleanUpPartialBootstrap()
        Debug.Print "Bootstrapper: Очистка после неудачной попытки..."
        If Not g_Logger Is Nothing Then g_Logger.Info "Bootstrapper: Начало очистки."
        Set g_DataManager = Nothing
        Set g_Config = Nothing
        ' Логгер обычно оставляют до последнего или используют специальный аварийный логгер
        ' Set g_Logger = Nothing 
        Debug.Print "Bootstrapper: Очистка завершена."
    End Sub

    Public Function IsBootstrapped() As Boolean
        IsBootstrapped = m_IsBootstrapped
    End Function
    
    ' --- Вспомогательные классы (заглушки для примера) ---
    ' CLogger - Модуль класса
    ' Public Sub Initialize(logPath As String): Debug.Print "Logger init: " & logPath: End Sub
    ' Public Sub Info(msg As String): Debug.Print "INFO: " & msg: End Sub
    ' Public Sub Error(msg As String): Debug.Print "ERROR: " & msg: End Sub

    ' CConfig - Модуль класса
    ' Private m_Settings As Object ' Scripting.Dictionary
    ' Public Sub LoadFromFile(configPath As String): Set m_Settings = CreateObject("Scripting.Dictionary"): m_Settings("DatabaseConnectionString") = "DSN=MyDB": Debug.Print "Config loaded from: " & configPath: End Sub
    ' Public Function GetSetting(key As String) As String: If m_Settings.Exists(key) Then GetSetting = m_Settings(key) Else GetSetting = "": End If: End Function

    ' CDataManager - Модуль класса
    ' Public Sub Initialize(connectionString As String): Debug.Print "DataManager init with: " & connectionString: End Sub
    ````
    ````vba
    ' В ThisWorkbook
    Private Sub Workbook_Open()
        AppBootstrapper.Execute Me
        
        ' Проверка успешности загрузки перед выполнением других действий
        If Not AppBootstrapper.IsBootstrapped() Then
            MsgBox "Приложение не было корректно инициализировано. Некоторые функции могут быть недоступны или работать некорректно.", vbExclamation
            ' Возможно, стоит закрыть книгу, если инициализация критична
            ' Me.Close False
        End If
    End Sub

    Private Sub Workbook_BeforeClose(Cancel As Boolean)
        If AppBootstrapper.IsBootstrapped() Then
            ' Вызов процедуры деинициализации, если она есть в Bootstrapper
            ' AppBootstrapper.ShutdownApplication
            AppBootstrapper.CleanUpPartialBootstrap ' В данном примере это очистка
        End If
    End Sub
    ````
*   **Резюме:** Bootstrapper централизует и упорядочивает процесс инициализации, делая его более управляемым, надежным и расширяемым. Он служит единой точкой для настройки всех основных компонентов приложения.
*   **Контрольные правила:**
    *   `Must`: Использовать Bootstrapper для приложений средней и высокой сложности.
    *   `Must`: Определить четкую последовательность инициализации компонентов в Bootstrapper.
    *   `Must`: Реализовать обработку ошибок на этапе начальной загрузки.
    *   `Must Not`: Дублировать логику инициализации вне Bootstrapper после его внедрения.

### 2.2. Подраздел: Инициализация глобальных и разделяемых ресурсов (Singleton)

*   **Проблема:** Необходимость обеспечить единственный экземпляр определенного класса (например, логгер, менеджер конфигурации, фабрика объектов) в масштабе всего приложения и предоставить к нему глобальную точку доступа.
*   **Причина (в контексте VBA):**
    *   VBA не имеет встроенной поддержки статических классов или свойств класса, как в C# или Java.
    *   Глобальные переменные в стандартных модулях могут быть сброшены при ошибках или при сбросе проекта, что приводит к потере экземпляра.
    *   Необходим контролируемый способ создания и доступа к единственному экземпляру.
*   **Ошибка (специфический симптом/пример):**
    *   Создание нескольких экземпляров класса, который должен быть единственным, что приводит к рассинхронизации состояния, избыточному потреблению ресурсов (например, несколько логгеров пишут в один файл с конфликтами).
    *   Использование `Public` переменной в стандартном модуле для хранения экземпляра без защиты от повторного создания или доступа до инициализации.
    ````vba
    ' Плохой подход к "глобальному" объекту
    ' Module: GlobalObjects
    Public MyGlobalLogger As CLogger ' Не CLogger, а например, CAppLogger
    
    Sub InitLogger()
        Set MyGlobalLogger = New CAppLogger ' Может быть вызвано несколько раз
    End Sub
    ````
*   **Решение (VBA-специфический паттерн/подход):** Адаптированный шаблон Singleton. Экземпляр класса создается при первом обращении через публичную функцию-аксессор в стандартном модуле. Эта функция проверяет, был ли уже создан экземпляр (хранящийся в `Private Static` переменной внутри функции или `Private` переменной уровня модуля), и если нет, то создает его.
    *   **Важно:** В VBA `Static` переменная в процедуре сохраняет свое значение между вызовами этой процедуры, но сбрасывается при сбросе VBA проекта (например, при ошибке времени выполнения, которая не обработана, или при нажатии кнопки "Стоп" в отладчике, или при изменении кода). Для большей устойчивости к сбросу состояния проекта (кроме полного закрытия/открытия хост-приложения или необработанных ошибок, приводящих к сбросу VBA) можно использовать переменную уровня модуля. Однако, для истинной персистентности между сессиями нужны внешние хранилища.
    *   Для VBA предпочтительнее использовать `Private` переменную уровня модуля для хранения экземпляра Singleton и публичную функцию `GetInstance` для доступа.

    **Реализация Singleton в VBA:**
    1.  **Модуль класса (например, `CApplicationSettings`)**:
        *   Конструктор (`Class_Initialize`) должен быть `Private`. Этого нельзя сделать напрямую в VBA. Поэтому создание контролируется исключительно через модуль-аксессор.
        *   Сам класс не должен иметь публичных методов для своего создания.
    2.  **Стандартный модуль (например, `SettingsService`)**:
        *   `Private` переменная уровня модуля для хранения единственного экземпляра.
        *   `Public Function GetInstance()`: проверяет, существует ли экземпляр. Если нет – создает его и сохраняет в `Private` переменной. Возвращает экземпляр.

*   **Примеры кода:**

    **Плохо (Bad):** Неконтролируемое создание "глобального" объекта.
    ````vba
    ' CAppConfig - Модуль класса
    ' Option Explicit
    ' Public ConfigValue As String
    ' Private Sub Class_Initialize()
    '     ConfigValue = "Default"
    '     Debug.Print "CAppConfig Initialized"
    ' End Sub

    ' Module1 - Стандартный модуль
    Public g_Config As CAppConfig ' Открыт для перезаписи извне

    Sub InitializeApp()
        Set g_Config = New CAppConfig ' Может быть вызвано много раз, создавая новые объекты
        g_Config.ConfigValue = "AppValue"
        Debug.Print "App Config Value: " & g_Config.ConfigValue
    End Sub

    Sub AnotherRoutine()
        ' Если InitializeApp не вызывался, g_Config is Nothing
        ' Если вызывался, используется существующий g_Config
        ' Но кто-то может сделать: Set Module1.g_Config = New CAppConfig, создав новый экземпляр
        If g_Config Is Nothing Then
            Debug.Print "Config не инициализирован в AnotherRoutine"
            Exit Sub
        End If
        Debug.Print "AnotherRoutine Config Value: " & g_Config.ConfigValue
    End Sub
    ````

    **Хорошо (Good):** Реализация Singleton.
    ````vba
    ' CApplicationSettings - Модуль класса (наш Singleton-объект)
    Option Explicit
    Private p_Settings As Object ' Scripting.Dictionary
    
    Private Sub Class_Initialize()
        Set p_Settings = CreateObject("Scripting.Dictionary")
        ' Загрузка настроек по умолчанию или из источника
        p_Settings("AppName") = "Мое Супер Приложение"
        p_Settings("Version") = "1.0"
        Debug.Print "Экземпляр CApplicationSettings создан и инициализирован."
    End Sub
    
    Public Function GetSetting(key As String) As Variant
        If p_Settings.Exists(key) Then
            GetSetting = p_Settings(key)
        Else
            GetSetting = CVErr(xlErrNA) ' Или vbNullString, или специфическое значение ошибки
        End If
    End Function
    
    Public Sub SetSetting(key As String, value As Variant)
        p_Settings(key) = value
    End Sub

    ' Friend Sub Terminate() ' Метод для явной очистки, если нужно
    ' Set p_Settings = Nothing
    ' Debug.Print "CApplicationSettings Terminated by friend call"
    ' End Sub
    
    Private Sub Class_Terminate()
        Set p_Settings = Nothing
        Debug.Print "Экземпляр CApplicationSettings уничтожен."
    End Sub

    ' AppSettingsService - Стандартный модуль (фабрика/аксессор для Singleton)
    Option Explicit
    Private m_oSettingsInstance As CApplicationSettings

    Public Function GetSettings() As CApplicationSettings
        If m_oSettingsInstance Is Nothing Then
            Set m_oSettingsInstance = New CApplicationSettings
        End If
        Set GetSettings = m_oSettingsInstance
    End Function
    
    ' Опционально: процедура для сброса Singleton (для тестов или специфических нужд)
    Public Sub ResetSettingsInstance()
        Set m_oSettingsInstance = Nothing
        Debug.Print "Экземпляр CApplicationSettings сброшен через AppSettingsService."
    End Sub

    ' Пример использования:
    ' Sub TestSingleton()
    '     Dim settings1 As CApplicationSettings
    '     Dim settings2 As CApplicationSettings
        
    '     Set settings1 = AppSettingsService.GetSettings()
    '     settings1.SetSetting "User", "Developer1"
        
    '     Set settings2 = AppSettingsService.GetSettings() ' Получит тот же самый экземпляр
    '     Debug.Print "User from settings2: " & settings2.GetSetting("User") ' Выведет "Developer1"
        
    '     Debug.Print "AppName: " & AppSettingsService.GetSettings().GetSetting("AppName")

    '     ' Проверка, что это один и тот же объект
    '     Debug.Print "settings1 Is settings2: " & (settings1 Is settings2) ' True
        
    '     ' Сброс для демонстрации (обычно не делается в рабочем коде без причины)
    '     ' AppSettingsService.ResetSettingsInstance
    '     ' Set settings1 = AppSettingsService.GetSettings() ' Будет создан новый экземпляр
    '     ' Debug.Print "After Reset, User from settings1: " & settings1.GetSetting("User") ' Будет пусто или значение по умолчанию
    ' End Sub
    ````
*   **Резюме:** Шаблон Singleton в VBA обеспечивает контролируемое создание единственного экземпляра класса и глобальный доступ к нему. Это полезно для управления общими ресурсами, такими как конфигурация или службы логирования. Реализация включает класс с приватной логикой инициализации и стандартный модуль с функцией-аксессором.
*   **Контрольные правила:**
    *   `Must`: Использовать шаблон Singleton для классов, которые должны иметь только один экземпляр в приложении.
    *   `Must`: Предоставлять доступ к экземпляру Singleton через публичную функцию-аксессор в стандартном модуле.
    *   `Must`: Хранить ссылку на экземпляр в `Private` переменной уровня модуля в модуле-аксессоре.
    *   `Must Not`: Предоставлять публичные конструкторы или методы для прямого создания экземпляра Singleton-класса вне модуля-аксессора.
    *   `Must Not`: Использовать глобальные переменные в стандартных модулях для хранения экземпляров Singleton без механизма контроля создания из `GetInstance`.

### 2.3. Подраздел: Отложенная (Lazy) инициализация

*   **Проблема:** Преждевременное создание и инициализация ресурсоемких объектов или компонентов, которые могут не понадобиться в текущем сеансе работы пользователя или понадобятся значительно позже.
*   **Причина (в контексте VBA):** Стремление инициализировать все компоненты приложения "заранее" при запуске (например, в `Workbook_Open` или Bootstrapper) может привести к замедлению старта приложения и излишнему потреблению памяти, если некоторые из этих компонентов используются редко или только при определенных условиях.
*   **Ошибка (специфический симптом/пример):**
    *   Медленный запуск приложения Excel, так как инициализируются все возможные отчеты, анализаторы данных, внешние подключения, даже если пользователь откроет книгу только для просмотра одного простого листа.
    *   Потребление памяти объектами, которые так и не были использованы.
*   **Решение (VBA-специфический паттерн/подход):** Инициализировать объект только в момент первого фактического обращения к нему. Это обычно реализуется через свойство `Get` или функцию-аксессор, которая проверяет, был ли объект уже создан. Если нет – создает и возвращает его; если да – просто возвращает существующий экземпляр.
*   **Примеры кода:**

    **Плохо (Bad):** Инициализация всего сразу в Bootstrapper.
    ````vba
    ' CHeavyReportGenerator - Модуль класса (ресурсоемкий)
    ' Option Explicit
    ' Private m_DataCache As Object ' Scripting.Dictionary
    ' Private Sub Class_Initialize()
    '     Set m_DataCache = CreateObject("Scripting.Dictionary")
    '     ' Имитация долгой загрузки данных
    '     Application.StatusBar = "Загрузка данных для генератора отчетов..."
    '     Dim i As Long
    '     For i = 1 To 500000 ' Имитация задержки
    '         If i Mod 10000 = 0 Then DoEvents
    '     Next i
    '     m_DataCache("SampleData") = "Очень много данных"
    '     Application.StatusBar = False
    '     Debug.Print "CHeavyReportGenerator инициализирован (загрузил данные)."
    ' End Sub
    ' Public Sub GenerateReport()
    '     Debug.Print "Отчет сгенерирован с использованием: " & m_DataCache("SampleData")
    ' End Sub

    ' AppServices - Стандартный модуль
    Public g_ReportGenerator As CHeavyReportGenerator

    Sub EagerInitialize() ' Вызывается, например, из Workbook_Open
        Debug.Print "Начало нетерпеливой инициализации..."
        Set g_ReportGenerator = New CHeavyReportGenerator ' Инициализация происходит сразу
        Debug.Print "Нетерпеливая инициализация завершена."
        ' Пользователь может никогда не вызвать g_ReportGenerator.GenerateReport
    End Sub
    ````

    **Хорошо (Good):** Отложенная инициализация через свойство доступа.
    ````vba
    ' CHeavyReportGenerator - Модуль класса (тот же, что и выше)
    ' Option Explicit
    ' Private m_DataCache As Object ' Scripting.Dictionary
    ' Private Sub Class_Initialize()
    '     Set m_DataCache = CreateObject("Scripting.Dictionary")
    '     Application.StatusBar = "Загрузка данных для генератора отчетов..."
    '     Dim i As Long
    '     For i = 1 To 500000 ' Имитация задержки
    '         If i Mod 10000 = 0 Then DoEvents
    '     Next i
    '     m_DataCache("SampleData") = "Очень много данных"
    '     Application.StatusBar = False
    '     Debug.Print "CHeavyReportGenerator инициализирован (загрузил данные)."
    ' End Sub
    ' Public Sub GenerateReport()
    '     Debug.Print "Отчет сгенерирован с использованием: " & m_DataCache("SampleData")
    ' End Sub
    ' Private Sub Class_Terminate()
    '     Set m_DataCache = Nothing
    '     Debug.Print "CHeavyReportGenerator деинициализирован."
    ' End Sub

    ' AppServices - Стандартный модуль
    Private m_LazyReportGenerator As CHeavyReportGenerator

    Public Property Get ReportGenerator() As CHeavyReportGenerator
        If m_LazyReportGenerator Is Nothing Then
            Debug.Print "ReportGenerator: Экземпляр еще не создан. Создание..."
            Set m_LazyReportGenerator = New CHeavyReportGenerator
        Else
            Debug.Print "ReportGenerator: Возврат существующего экземпляра."
        End If
        Set ReportGenerator = m_LazyReportGenerator
    End Property
    
    ' Опционально: метод для явного освобождения ресурса, если он больше не нужен
    Public Sub ReleaseReportGenerator()
        If Not m_LazyReportGenerator Is Nothing Then
            Debug.Print "ReportGenerator: Явное освобождение экземпляра."
            Set m_LazyReportGenerator = Nothing
        End If
    End Sub

    ' Пример использования:
    ' Sub TestLazyInitialization()
    '     Debug.Print "Тест отложенной инициализации начат."
    '     ' На данный момент CHeavyReportGenerator еще не создан
        
    '     Dim choice As VbMsgBoxResult
    '     choice = MsgBox("Сгенерировать тяжелый отчет?", vbYesNo)
        
    '     If choice = vbYes Then
    '         ' Только сейчас, при первом обращении к AppServices.ReportGenerator,
    '         ' будет создан экземпляр CHeavyReportGenerator
    '         AppServices.ReportGenerator.GenerateReport
    '         MsgBox "Отчет сгенерирован!"
    '     Else
    '         MsgBox "Генерация отчета отменена. Ресурсоемкий объект не был создан."
    '     End If
        
    '     ' Если нужно освободить память (например, перед закрытием части функционала)
    '     ' AppServices.ReleaseReportGenerator
    '     Debug.Print "Тест отложенной инициализации завершен."
    ' End Sub
    ````
*   **Резюме:** Отложенная инициализация улучшает производительность запуска приложения и экономит ресурсы, создавая объекты только тогда, когда они действительно необходимы. Это особенно полезно для компонентов, которые являются ресурсоемкими или используются нечасто.
*   **Контрольные правила:**
    *   `Must`: Рассмотреть использование отложенной инициализации для ресурсоемких или редко используемых объектов.
    *   `Must`: Реализовывать отложенную инициализацию через свойство `Get` или функцию-аксессор, которая инкапсулирует логику создания объекта "по требованию".
    *   `Must Not`: Инициализировать все объекты приложения при запуске, если некоторые из них могут не использоваться.

### Резюме Раздела 2
Стратегии инициализации, такие как Bootstrapping, Singleton и Lazy Initialization, помогают структурировать запуск VBA-приложения, управлять глобальными ресурсами и оптимизировать производительность. Bootstrapper координирует общий процесс, Singleton обеспечивает уникальность экземпляров критически важных сервисов, а Lazy Initialization откладывает создание ресурсоемких объектов до момента их реального использования. Эти подходы делают приложение более надежным, управляемым и эффективным.

### Золотые правила Раздела 2
1.  **Централизуй запуск:** Используй Bootstrapper для координации инициализации сложных приложений.
2.  **Один для всех (где нужно):** Применяй Singleton для ресурсов, требующих единственного экземпляра (конфигурация, логгер).
3.  **Не спеши (если можно):** Используй Lazy Initialization для тяжелых или редко используемых объектов для ускорения старта и экономии ресурсов.
4.  **Управляй зависимостями:** Учитывай порядок инициализации компонентов; компоненты с меньшим количеством зависимостей или без них должны инициализироваться первыми.

### Мастер-класс Раздела 2: Приложение с Bootstrapper, Singleton (Логгер) и Lazy Initialization (Анализатор Данных)

**Сценарий:** Приложение Excel, которое при запуске инициализирует логгер (Singleton). Основная функциональность – анализ данных на листе "InputData" – выполняется "Анализатором Данных", который является ресурсоемким и должен инициализироваться отложенно.

**1. Модуль класса `CAppLogger` (Singleton)**
````vba
' CAppLogger - Модуль класса
Option Explicit
Private m_LogFilePath As String
Private m_FSO As Object ' FileSystemObject
Private m_LogStream As Object ' TextStream

Private Sub Class_Initialize()
    Set m_FSO = CreateObject("Scripting.FileSystemObject")
    Debug.Print "CAppLogger: Экземпляр создан."
End Sub

Public Sub InitializeLog(ByVal logFilePath As String)
    m_LogFilePath = logFilePath
    On Error Resume Next ' Подавить ошибку, если файл уже открыт другим процессом в монопольном режиме
    Set m_LogStream = m_FSO.OpenTextFile(m_LogFilePath, 8, True, -1) ' 8=Append, True=Create, -1=Unicode
    If Err.Number <> 0 Then
        Debug.Print "CAppLogger: Не удалось открыть файл лога " & m_LogFilePath & ". Ошибка: " & Err.Description
        Set m_LogStream = Nothing
    Else
        Debug.Print "CAppLogger: Логгирование в файл " & m_LogFilePath
    End If
    On Error GoTo 0
End Sub

Public Sub Log(ByVal message As String, Optional ByVal logLevel As String = "INFO")
    Dim logEntry As String
    logEntry = Now() & " [" & logLevel & "] - " & message
    Debug.Print logEntry ' Всегда выводим в Immediate Window
    If Not m_LogStream Is Nothing Then
        On Error Resume Next ' Если файл стал недоступен
        m_LogStream.WriteLine logEntry
        If Err.Number <> 0 Then
            Debug.Print "CAppLogger: Ошибка записи в лог-файл. " & Err.Description
            ' Можно попытаться переоткрыть файл или просто пропустить запись
        End If
        On Error GoTo 0
    End If
End Sub

Private Sub Class_Terminate()
    If Not m_LogStream Is Nothing Then
        m_LogStream.Close
        Set m_LogStream = Nothing
    End If
    Set m_FSO = Nothing
    Debug.Print "CAppLogger: Экземпляр уничтожен, лог-файл закрыт."
End Sub
````

**2. Стандартный модуль `LoggerService` (для Singleton CAppLogger)**
````vba
' LoggerService - Стандартный модуль
Option Explicit
Private s_LoggerInstance As CAppLogger

Public Function GetLogger() As CAppLogger
    If s_LoggerInstance Is Nothing Then
        Set s_LoggerInstance = New CAppLogger
    End If
    Set GetLogger = s_LoggerInstance
End Function

Public Sub ReleaseLogger() ' Для явного освобождения, если потребуется
    Set s_LoggerInstance = Nothing
End Sub
````

**3. Модуль класса `CDataAnalyzer` (Ресурсоемкий, для Lazy Init)**
````vba
' CDataAnalyzer - Модуль класса
Option Explicit
Private m_DataSourceSheet As Worksheet
Private m_AnalysisResults As Object ' Scripting.Dictionary

Private Sub Class_Initialize()
    LoggerService.GetLogger.Log "CDataAnalyzer: Начало инициализации..."
    Set m_AnalysisResults = CreateObject("Scripting.Dictionary")
    ' Имитация длительной загрузки/подготовки
    Dim i As Long
    Application.StatusBar = "CDataAnalyzer: Подготовка данных..."
    For i = 1 To 300000 ' Уменьшим для скорости демонстрации
        If i Mod 10000 = 0 Then DoEvents
    Next i
    Application.StatusBar = False
    LoggerService.GetLogger.Log "CDataAnalyzer: Экземпляр создан и готов."
End Sub

Public Sub Configure(ByVal dataSource As Worksheet)
    Set m_DataSourceSheet = dataSource
    LoggerService.GetLogger.Log "CDataAnalyzer: Сконфигурирован для листа " & dataSource.Name
End Sub

Public Function PerformAnalysis() As String
    If m_DataSourceSheet Is Nothing Then
        PerformAnalysis = "Ошибка: Источник данных не настроен."
        LoggerService.GetLogger.Log "CDataAnalyzer: Анализ невозможен, источник не настроен.", "ERROR"
        Exit Function
    End If
    
    LoggerService.GetLogger.Log "CDataAnalyzer: Начало анализа данных на листе " & m_DataSourceSheet.Name
    ' Имитация анализа
    Dim rowCount As Long
    rowCount = m_DataSourceSheet.Cells(Rows.Count, "A").End(xlUp).Row
    m_AnalysisResults("TotalRows") = rowCount
    m_AnalysisResults("AnalysisTimestamp") = Now()
    
    Dim i As Long
    Application.StatusBar = "CDataAnalyzer: Выполнение анализа..."
    For i = 1 To 200000 ' Имитация задержки
        If i Mod 10000 = 0 Then DoEvents
    Next i
    Application.StatusBar = False
    
    PerformAnalysis = "Анализ завершен. Обработано строк: " & rowCount & ". Время: " & Now()
    LoggerService.GetLogger.Log "CDataAnalyzer: " & PerformAnalysis, "SUCCESS"
End Function

Private Sub Class_Terminate()
    Set m_DataSourceSheet = Nothing
    Set m_AnalysisResults = Nothing
    LoggerService.GetLogger.Log "CDataAnalyzer: Экземпляр уничтожен."
End Sub
````

**4. Стандартный модуль `AppServices` (для Lazy Init CDataAnalyzer)**
````vba
' AppServices - Стандартный модуль
Option Explicit
Private s_DataAnalyzerInstance As CDataAnalyzer

Public Property Get DataAnalyzer() As CDataAnalyzer
    LoggerService.GetLogger.Log "AppServices: Запрос экземпляра DataAnalyzer..."
    If s_DataAnalyzerInstance Is Nothing Then
        LoggerService.GetLogger.Log "AppServices: DataAnalyzer еще не создан. Создание..."
        Set s_DataAnalyzerInstance = New CDataAnalyzer
    End If
    Set DataAnalyzer = s_DataAnalyzerInstance
End Property

Public Sub ReleaseDataAnalyzer()
    If Not s_DataAnalyzerInstance Is Nothing Then
        LoggerService.GetLogger.Log "AppServices: Явное освобождение DataAnalyzer."
        Set s_DataAnalyzerInstance = Nothing
    End If
End Sub
````

**5. Стандартный модуль `AppBootstrapper`**
````vba
' AppBootstrapper - Стандартный модуль
Option Explicit
Private m_IsBootstrapped As Boolean

Public Sub Execute(ByVal wb As Workbook)
    If m_IsBootstrapped Then Exit Sub

    Debug.Print "AppBootstrapper: Запуск..."
    
    ' 1. Инициализация Логгера (Singleton)
    LoggerService.GetLogger.InitializeLog wb.Path & "\MasterclassApp.log"
    LoggerService.GetLogger.Log "AppBootstrapper: Логгер инициализирован."
    
    ' 2. Другие немедленные инициализации (если есть)
    ' Например, загрузка основной конфигурации (не показано в этом примере)
    LoggerService.GetLogger.Log "AppBootstrapper: Основные сервисы настроены."

    m_IsBootstrapped = True
    LoggerService.GetLogger.Log "AppBootstrapper: Успешно завершен."
    MsgBox "Приложение инициализировано (Логгер активен). Анализатор данных будет создан по требованию.", vbInformation
End Sub

Public Sub Shutdown()
    If Not m_IsBootstrapped Then Exit Sub
    LoggerService.GetLogger.Log "AppBootstrapper: Начало процедуры Shutdown..."
    
    ' Освобождение отложенно инициализируемых ресурсов
    AppServices.ReleaseDataAnalyzer
    
    ' Освобождение Singleton ресурсов
    LoggerService.ReleaseLogger ' Важно, чтобы логгер освобождался последним или корректно обрабатывал логирование во время Shutdown
    
    m_IsBootstrapped = False
    Debug.Print "AppBootstrapper: Shutdown завершен." ' Логгер уже может быть недоступен
End Sub

Public Function IsBootstrapped() As Boolean
    IsBootstrapped = m_IsBootstrapped
End Function
````

**6. Модуль `ThisWorkbook`**
````vba
' ThisWorkbook - Модуль книги
Option Explicit

Private Sub Workbook_Open()
    AppBootstrapper.Execute Me
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    AppBootstrapper.Shutdown
End Sub
````

**7. Пример использования (например, кнопка на листе вызывает этот макрос)**
````vba
' Module: UserActions (Стандартный модуль)
Option Explicit

Sub PerformDataAnalysisAction()
    If Not AppBootstrapper.IsBootstrapped() Then
        MsgBox "Приложение не инициализировано!", vbCritical
        Exit Sub
    End If

    LoggerService.GetLogger.Log "UserActions: Запрошено выполнение анализа данных."
    
    Dim analyzer As CDataAnalyzer
    Dim targetSheet As Worksheet
    
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Worksheets("InputData")
    On Error GoTo 0
    
    If targetSheet Is Nothing Then
        MsgBox "Лист 'InputData' не найден! Невозможно выполнить анализ.", vbExclamation
        LoggerService.GetLogger.Log "UserActions: Лист 'InputData' не найден.", "WARN"
        Exit Sub
    End If
    
    ' Получение экземпляра анализатора (будет создан при первом обращении здесь)
    Set analyzer = AppServices.DataAnalyzer
    analyzer.Configure targetSheet ' Настройка источника данных
    
    Dim result As String
    result = analyzer.PerformAnalysis()
    
    MsgBox result, vbInformation, "Результат анализа"
    LoggerService.GetLogger.Log "UserActions: Анализ данных завершен. Результат показан пользователю."
End Sub
````

**Как это работает:**
1.  `Workbook_Open` вызывает `AppBootstrapper.Execute`.
2.  Bootstrapper инициализирует логгер (`CAppLogger` через `LoggerService.GetLogger()`). `CAppLogger` создается как Singleton. Лог-файл начинает записываться.
3.  Bootstrapper завершает свою работу. `CDataAnalyzer` еще **не создан**.
4.  Пользователь нажимает кнопку, которая вызывает `UserActions.PerformDataAnalysisAction`.
5.  Процедура запрашивает `AppServices.DataAnalyzer`.
6.  Свойство `AppServices.DataAnalyzer` (Property Get) видит, что `s_DataAnalyzerInstance` равен `Nothing`, создает новый экземпляр `CDataAnalyzer` (в этот момент происходит "ресурсоемкая" инициализация `CDataAnalyzer`, и он пишет в лог через `LoggerService.GetLogger()`), и возвращает его.
7.  Анализатор настраивается и выполняет анализ. Все действия логируются.
8.  При закрытии книги `Workbook_BeforeClose` вызывает `AppBootstrapper.Shutdown`, который освобождает ресурсы, включая `DataAnalyzer` и `Logger`.

Этот мастер-класс демонстрирует, как Bootstrapper управляет запуском, Singleton обеспечивает единый логгер, а Lazy Initialization откладывает создание `CDataAnalyzer` до момента, когда он действительно понадобится, оптимизируя запуск и использование ресурсов.

---
## Раздел 3: Организация доступа к функциональности

### 3.1. Подраздел: Использование стандартных модулей vs. модулей классов

*   **Проблема:** Неправильный выбор между стандартными модулями и модулями классов для размещения кода, что приводит к плохо структурированным, трудно поддерживаемым и тестируемым приложениям.
*   **Причина (в контексте VBA):**
    *   **Стандартные модули (.bas):** Содержат процедуры (`Sub`) и функции (`Function`), которые по своей природе похожи на статические методы в других языках. Они не могут быть инстанциированы (нельзя написать `Dim mod As New Module1`). Переменные, объявленные как `Public` в стандартном модуле, имеют глобальную область видимости в пределах проекта VBA. Состояние, хранимое в переменных уровня модуля, является глобальным и может быть проблематичным для управления.
    *   **Модули классов (.cls):** Являются шаблонами для создания объектов. Каждый объект, созданный из класса (`Set obj = New MyClass`), имеет свой собственный набор переменных экземпляра (членов класса), что позволяет управлять состоянием инкапсулировано. Классы поддерживают свойства (`Property Get/Let/Set`) и методы.
*   **Ошибка (специфический симптом/пример):**
    *   **Чрезмерное использование стандартных модулей:** Вся или почти вся логика приложения, включая управление состоянием, размещается в стандартных модулях с использованием глобальных переменных. Это приводит к высокой связанности, трудностям в тестировании отдельных частей, риску конфликта имен и неконтролируемому изменению состояния из любой точки кода.
        ````vba
        ' Module: BadGlobalStateManager
        Public g_CurrentUserName As String
        Public g_LastProcessedID As Long
        Public g_DataArray() As Variant ' Глобальный массив данных

        Sub InitializeUserState(userName As String)
            g_CurrentUserName = userName
            g_LastProcessedID = 0
            ' Загрузка данных в g_DataArray...
            Debug.Print "Пользователь " & g_CurrentUserName & " инициализирован глобально."
        End Sub

        Sub ProcessNextItem()
            If g_CurrentUserName = "" Then Exit Sub
            g_LastProcessedID = g_LastProcessedID + 1
            ' Обработка g_DataArray(g_LastProcessedID)...
            Debug.Print "Обработан элемент " & g_LastProcessedID & " для " & g_CurrentUserName
        End Sub
        ' Проблема: g_CurrentUserName, g_LastProcessedID, g_DataArray могут быть изменены откуда угодно.
        ' Если нужно два "сеанса" обработки параллельно - невозможно без усложнений.
        ````
    *   **Избыточное использование классов для простой утилитарной логики:** Создание классов для функций, которые не имеют состояния и могли бы быть простыми функциями в стандартном модуле (например, класс "MathHelpers" с одним методом "Add"). Это добавляет ненужную сложность создания экземпляра.
*   **Решение (VBA-специфический паттерн/подход):**
    *   **Стандартные модули использовать для:**
        *   **Утилитарных функций:** Функции, которые не зависят от состояния или работают исключительно с переданными им аргументами (например, математические расчеты, форматирование строк, общие операции с файлами без сохранения состояния между вызовами).
        *   **Точек входа приложения:** `Workbook_Open`, процедуры, вызываемые с ленты или кнопок, которые делегируют работу объектам.
        *   **Процедур Bootstrapping/Startup.**
        *   **Фабрик или сервисов доступа к Singleton'ам** (как `LoggerService.GetLogger()`).
        *   Глобальных констант (`Public Const`).
    *   **Модули классов использовать для:**
        *   **Представления сущностей и бизнес-объектов:** Объекты, которые имеют состояние и поведение (например, `CEmployee`, `CInvoice`, `CProduct`).
        *   **Сервисных объектов:** Компоненты, инкапсулирующие сложную логику или управляющие ресурсами (например, `CDataManager`, `CReportGenerator`, `CExternalAPIClient`).
        *   **Реализации паттернов проектирования:** Таких как Strategy, Observer, Command, где требуется полиморфизм или инкапсуляция поведения.
        *   **Управления UI элементами:** Например, класс для управления сложной пользовательской формой или набором элементов управления на листе.
*   **Примеры кода:**

    **Плохо (Bad):** Вся логика в стандартном модуле.
    ````vba
    ' Module: AllInOneModule
    Public CurrentFile As String
    Public DataLoaded As Boolean
    Public ProcessedData As Object ' Scripting.Dictionary

    Sub LoadDataFromFile(filePath As String)
        ' ... логика загрузки ...
        CurrentFile = filePath
        Set ProcessedData = CreateObject("Scripting.Dictionary")
        ' ... заполнение ProcessedData ...
        DataLoaded = True
        Debug.Print "Данные загружены из " & CurrentFile
    End Sub

    Sub FilterData(filterCriteria As String)
        If Not DataLoaded Then Exit Sub
        ' ... логика фильтрации ProcessedData ...
        Debug.Print "Данные отфильтрованы по: " & filterCriteria
    End Sub

    Sub SaveFilteredData(newFilePath As String)
        If Not DataLoaded Then Exit Sub
        ' ... логика сохранения отфильтрованных данных из ProcessedData ...
        Debug.Print "Отфильтрованные данные сохранены в " & newFilePath
    End Sub
    ' Проблемы: состояние (CurrentFile, DataLoaded, ProcessedData) глобально,
    ' трудно тестировать, невозможно иметь два "экземпляра" обработки файлов.
    ````

    **Хорошо (Good):** Разделение на класс и стандартный модуль.
    ````vba
    ' CFileProcessor - Модуль класса
    Option Explicit
    Private m_FilePath As String
    Private m_FileData As Object ' Scripting.Dictionary
    Private m_IsDataLoaded As Boolean

    Public Sub LoadData(ByVal filePath As String)
        m_FilePath = filePath
        Set m_FileData = CreateObject("Scripting.Dictionary")
        ' Имитация загрузки
        m_FileData("Row1") = "Sample Data Alpha"
        m_FileData("Row2") = "Sample Data Beta"
        m_IsDataLoaded = True
        Debug.Print "CFileProcessor: Данные загружены из " & m_FilePath
    End Sub

    Public Function FilterData(ByVal filterCriteria As String) As Object ' Scripting.Dictionary
        If Not m_IsDataLoaded Then
            Debug.Print "CFileProcessor: Данные не загружены для фильтрации."
            Set FilterData = Nothing
            Exit Function
        End If
        
        Dim filteredDict As Object ' Scripting.Dictionary
        Set filteredDict = CreateObject("Scripting.Dictionary")
        ' Имитация фильтрации
        Dim key As Variant
        For Each key In m_FileData.Keys
            If InStr(1, m_FileData(key), filterCriteria, vbTextCompare) > 0 Then
                filteredDict(key) = m_FileData(key)
            End If
        Next key
        
        Debug.Print "CFileProcessor: Данные отфильтрованы по '" & filterCriteria & "'. Найдено: " & filteredDict.Count
        Set FilterData = filteredDict
    End Function

    Public Sub SaveData(ByVal dataToSave As Object, ByVal newFilePath As String)
        If dataToSave Is Nothing Then
            Debug.Print "CFileProcessor: Нет данных для сохранения."
            Exit Sub
        End If
        ' Имитация сохранения
        Debug.Print "CFileProcessor: Данные сохранены в " & newFilePath & ". Количество элементов: " & dataToSave.Count
    End Sub

    Private Sub Class_Initialize()
        Debug.Print "CFileProcessor: Экземпляр создан."
        m_IsDataLoaded = False
    End Sub

    Private Sub Class_Terminate()
        Set m_FileData = Nothing
        Debug.Print "CFileProcessor: Экземпляр уничтожен."
    End Sub

    ' FileUtils - Стандартный модуль (для утилит и точки вызова)
    Option Explicit
    Public Function CheckFileExists(ByVal filePath As String) As Boolean
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        CheckFileExists = fso.FileExists(filePath)
        Set fso = Nothing
    End Function

    Sub ProcessMyFile(originalFile As String, filter As String, outputFile As String)
        If Not CheckFileExists(originalFile) Then
            MsgBox "Файл не найден: " & originalFile, vbCritical
            Exit Sub
        End If

        Dim processor As CFileProcessor
        Set processor = New CFileProcessor
        
        processor.LoadData originalFile
        
        Dim filtered As Object
        Set filtered = processor.FilterData(filter)
        
        If Not filtered Is Nothing And filtered.Count > 0 Then
            processor.SaveData filtered, outputFile
            MsgBox "Обработка файла завершена. Результат в " & outputFile
        Else
            MsgBox "После фильтрации не осталось данных для сохранения.", vbInformation
        End If
        
        Set processor = Nothing ' Освобождаем объект
    End Sub

    ' Пример вызова:
    ' Sub TestFileProcessing()
    '     ' Создайте файл C:\temp\mydata.txt для теста
    '     Call ProcessMyFile("C:\temp\mydata.txt", "Alpha", "C:\temp\filtered_data.txt")
    ' End Sub
    ````
*   **Резюме:** Правильное распределение кода между стандартными модулями и модулями классов является ключом к созданию хорошо структурированных VBA-приложений. Стандартные модули подходят для утилит и процедур без состояния, в то время как классы инкапсулируют состояние и поведение, формируя многократно используемые компоненты.
*   **Контрольные правила:**
    *   `Must`: Использовать модули классов для представления сущностей, сервисов и любой логики, требующей управления состоянием или инкапсуляции.
    *   `Must`: Использовать стандартные модули для утилитарных функций, глобальных констант и процедур, координирующих работу объектов (например, Bootstrapper, точки входа).
    *   `Must Not`: Размещать сложную бизнес-логику с управлением состоянием в стандартных модулях с использованием множества глобальных переменных.
    *   `Must Not`: Создавать классы для простых, не имеющих состояния функций, если они могут быть эффективно реализованы в стандартном модуле.

### 3.2. Подраздел: Фасады для упрощения взаимодействия

*   **Проблема:** Клиентский код (код, использующий некоторую функциональность) становится слишком сложным и тесно связанным с множеством мелких объектов сложной подсистемы. Это затрудняет понимание, использование и изменение подсистемы.
*   **Причина (в контексте VBA):** Приложение может состоять из нескольких взаимодействующих классов, каждый из которых отвечает за свою часть задачи. Чтобы выполнить общую бизнес-операцию, клиентскому коду приходится знать обо всех этих классах, правильно их инициализировать и координировать их вызовы в нужной последовательности.
*   **Ошибка (специфический симптом/пример):** Процедура в стандартном модуле или метод в другом классе содержит десятки строк кода, создающих и вызывающих методы 3-5 различных объектов для выполнения одной, казалось бы, простой задачи (например, "Оформить заказ"). Любое изменение во внутренней структуре этой подсистемы оформления заказа (например, добавление нового шага проверки) потребует изменения во всех местах, где эта логика используется.
*   **Решение (VBA-специфический паттерн/подход):** Реализация шаблона "Фасад". Фасад – это класс, который предоставляет единый, упрощенный интерфейс к набору интерфейсов в подсистеме. Он инкапсулирует сложность взаимодействия между компонентами подсистемы и предоставляет клиентам высокоуровневые методы.
*   **Примеры кода:**

    **Плохо (Bad):** Клиент напрямую работает со сложной подсистемой.
    ````vba
    ' --- Компоненты подсистемы ---
    ' CInventoryManager - Модуль класса
    ' Public Function CheckStock(itemCode As String, quantity As Long) As Boolean: Debug.Print "Inventory: Checking " & itemCode; quantity: CheckStock = True: End Function
    ' Public Sub ReserveStock(itemCode As String, quantity As Long): Debug.Print "Inventory: Reserving " & itemCode; quantity: End Sub

    ' CPaymentProcessor - Модуль класса
    ' Public Function ProcessPayment(amount As Double, cardDetails As String) As Boolean: Debug.Print "Payment: Processing " & amount: ProcessPayment = True: End Function

    ' CNotificationService - Модуль класса
    ' Public Sub SendConfirmation(email As String, message As String): Debug.Print "Notification: Sending to " & email & ": " & message: End Sub

    ' --- Клиентский код (например, в модуле обработки UI) ---
    ' Sub PlaceOrderAction(customerEmail As String, item As String, qty As Long, price As Double, paymentInfo As String)
    '     Dim inventory As CInventoryManager
    '     Dim payment As CPaymentProcessor
    '     Dim notifier As CNotificationService
        
    '     Set inventory = New CInventoryManager
    '     Set payment = New CPaymentProcessor
    '     Set notifier = New CNotificationService
        
    '     Debug.Print "Клиент: Начало оформления заказа..."
    '     If Not inventory.CheckStock(item, qty) Then
    '         MsgBox "Товара нет на складе!", vbExclamation
    '         Set inventory = Nothing: Set payment = Nothing: Set notifier = Nothing
    '         Exit Sub
    '     End If
        
    '     inventory.ReserveStock item, qty
        
    '     If Not payment.ProcessPayment(qty * price, paymentInfo) Then
    '         MsgBox "Платеж не прошел!", vbCritical
    '         ' Нужна логика отката резервации товара (не показана)
    '         Set inventory = Nothing: Set payment = Nothing: Set notifier = Nothing
    '         Exit Sub
    '     End If
        
    '     notifier.SendConfirmation customerEmail, "Ваш заказ " & item & " (" & qty & " шт.) успешно оформлен."
    '     MsgBox "Заказ успешно оформлен!", vbInformation
        
    '     Set inventory = Nothing
    '     Set payment = Nothing
    '     Set notifier = Nothing
    '     Debug.Print "Клиент: Оформление заказа завершено."
    ' End Sub
    ' Проблема: Клиентский код знает все детали подсистемы. Много объектов для управления.
    ````

    **Хорошо (Good):** Использование Фасада.
    ````vba
    ' --- Компоненты подсистемы (те же CInventoryManager, CPaymentProcessor, CNotificationService) ---
    ' CInventoryManager - Модуль класса
    ' Option Explicit
    ' Public Function CheckStock(itemCode As String, quantity As Long) As Boolean: Debug.Print "Inventory: Checking " & itemCode; quantity: CheckStock = (itemCode <> "SOLD_OUT"): End Function
    ' Public Sub ReserveStock(itemCode As String, quantity As Long): Debug.Print "Inventory: Reserving " & itemCode; quantity: End Sub
    ' Public Sub ReleaseStock(itemCode As String, quantity As Long): Debug.Print "Inventory: Releasing " & itemCode; quantity: End Sub ' Для отката

    ' CPaymentProcessor - Модуль класса
    ' Option Explicit
    ' Public Function ProcessPayment(amount As Double, cardDetails As String) As Boolean: Debug.Print "Payment: Processing " & amount: ProcessPayment = (cardDetails <> "INVALID_CARD"): End Function

    ' CNotificationService - Модуль класса
    ' Option Explicit
    ' Public Sub SendConfirmation(email As String, message As String): Debug.Print "Notification: Sending to " & email & ": " & message: End Sub
    ' Public Sub SendFailureNotice(email As String, message As String): Debug.Print "Notification: Sending failure to " & email & ": " & message: End Sub


    ' COrderProcessingFacade - Модуль класса (Фасад)
    Option Explicit
    Private m_Inventory As CInventoryManager
    Private m_Payment As CPaymentProcessor
    Private m_Notifier As CNotificationService

    Private Sub Class_Initialize()
        Set m_Inventory = New CInventoryManager
        Set m_Payment = New CPaymentProcessor
        Set m_Notifier = New CNotificationService
        Debug.Print "COrderProcessingFacade: Инициализирован."
    End Sub

    Public Function PlaceOrder(customerEmail As String, itemCode As String, quantity As Long, unitPrice As Double, paymentDetails As String) As Boolean
        Dim success As Boolean
        success = False
        LoggerService.GetLogger.Log "Facade: Начало оформления заказа для " & customerEmail & ", товар: " & itemCode
        
        ' 1. Проверить наличие товара
        If Not m_Inventory.CheckStock(itemCode, quantity) Then
            LoggerService.GetLogger.Log "Facade: Товара " & itemCode & " нет на складе.", "WARN"
            m_Notifier.SendFailureNotice customerEmail, "Не удалось оформить заказ: товар " & itemCode & " отсутствует на складе."
            PlaceOrder = False
            Exit Function
        End If
        LoggerService.GetLogger.Log "Facade: Товар " & itemCode & " в наличии."

        ' 2. Зарезервировать товар
        m_Inventory.ReserveStock itemCode, quantity
        LoggerService.GetLogger.Log "Facade: Товар " & itemCode & " зарезервирован."

        ' 3. Обработать платеж
        If Not m_Payment.ProcessPayment(quantity * unitPrice, paymentDetails) Then
            LoggerService.GetLogger.Log "Facade: Платеж не прошел.", "ERROR"
            m_Inventory.ReleaseStock itemCode, quantity ' Откат резервации
            LoggerService.GetLogger.Log "Facade: Резерв товара " & itemCode & " отменен."
            m_Notifier.SendFailureNotice customerEmail, "Не удалось оформить заказ: проблема с оплатой."
            PlaceOrder = False
            Exit Function
        End If
        LoggerService.GetLogger.Log "Facade: Платеж успешно обработан."

        ' 4. Отправить подтверждение
        m_Notifier.SendConfirmation customerEmail, "Ваш заказ на " & itemCode & " (" & quantity & " шт.) успешно оформлен и оплачен."
        LoggerService.GetLogger.Log "Facade: Подтверждение заказа отправлено на " & customerEmail, "SUCCESS"
        
        success = True
        PlaceOrder = success
    End Function

    Private Sub Class_Terminate()
        Set m_Inventory = Nothing
        Set m_Payment = Nothing
        Set m_Notifier = Nothing
        Debug.Print "COrderProcessingFacade: Деинициализирован."
    End Sub

    ' --- Клиентский код (например, в модуле обработки UI или AppMain) ---
    ' Sub TestFacadeOrderPlacement()
    '     ' Убедитесь, что LoggerService и CAppLogger определены и Bootstrapper вызвал инициализацию логгера
    '     ' AppBootstrapper.Execute ThisWorkbook (если еще не вызван)
    '
    '     Dim orderFacade As COrderProcessingFacade
    '     Set orderFacade = New COrderProcessingFacade
    '
    '     Dim customer As String
    '     customer = "test@example.com"
    '
    '     ' Успешный заказ
    '     If orderFacade.PlaceOrder(customer, "ITEM001", 2, 10.50, "VALID_CARD_DETAILS") Then
    '         MsgBox "Заказ для " & customer & " успешно оформлен через фасад!", vbInformation
    '     Else
    '         MsgBox "Не удалось оформить заказ для " & customer & " через фасад.", vbExclamation
    '     End If
    '
    '     ' Неудачный заказ (нет на складе)
    '     If orderFacade.PlaceOrder(customer, "SOLD_OUT", 1, 25.00, "VALID_CARD_DETAILS") Then
    '         MsgBox "Заказ (SOLD_OUT) успешно оформлен через фасад!", vbInformation ' Не должно произойти
    '     Else
    '         MsgBox "Не удалось оформить заказ (SOLD_OUT) для " & customer & " через фасад.", vbExclamation
    '     End If
    '
    '     Set orderFacade = Nothing
    ' End Sub
    ````
*   **Резюме:** Шаблон Фасад упрощает взаимодействие со сложными подсистемами, предоставляя единый, высокоуровневый интерфейс. Это снижает связанность между клиентским кодом и компонентами подсистемы, облегчает использование и поддержку системы.
*   **Контрольные правила:**
    *   `Must`: Использовать Фасад для предоставления упрощенного доступа к сложной подсистеме из множества взаимодействующих классов.
    *   `Must`: Инкапсулировать логику координации компонентов подсистемы внутри методов Фасада.
    *   `Must Not`: Добавлять в Фасад новую бизнес-логику, которая не относится к координации существующих компонентов подсистемы. Фасад – это упрощенный "пульт управления", а не новый "двигатель".
    *   `Must Not`: Делать все классы в приложении доступными только через один гигантский Фасад (это может привести к созданию "Божественного объекта"). Используйте Фасады для логически сгруппированных подсистем.

### Резюме Раздела 3
Организация доступа к функциональности в VBA требует осмысленного выбора между стандартными модулями и модулями классов, а также применения паттернов, таких как Фасад, для упрощения взаимодействия со сложными частями системы. Правильное разделение ответственности и инкапсуляция сложности ведут к более чистой, гибкой и поддерживаемой архитектуре.

### Золотые правила Раздела 3
1.  **Классы для состояния, модули для утилит:** Используйте классы для объектов с состоянием и поведением; стандартные модули – для вспомогательных функций и координации.
2.  **Инкапсулируй сложность:** Скрывайте детали реализации сложных подсистем за Фасадами.
3.  **Слабая связанность, высокая сплоченность:** Стремитесь к тому, чтобы компоненты были как можно менее зависимы друг от друга (слабая связанность) и чтобы каждый компонент (класс или модуль) выполнял четко определенную, сфокусированную задачу (высокая сплоченность).
4.  **Ясные интерфейсы:** Определяйте понятные и минимально необходимые публичные интерфейсы для ваших классов и модулей.

### Мастер-класс Раздела 3: Система обработки заказов с Фасадом

**Сценарий:** Расширим предыдущий пример с Фасадом для обработки заказов. Добавим класс `CProductCatalog` и интегрируем его в `COrderProcessingFacade`. Клиентский код будет взаимодействовать только с Фасадом для получения информации о товаре и размещения заказа.

**1. Модуль класса `CProductCatalog`**
````vba
' CProductCatalog - Модуль класса
Option Explicit
Private m_Products As Object ' Scripting.Dictionary

Private Sub Class_Initialize()
    Set m_Products = CreateObject("Scripting.Dictionary")
    ' Загрузка каталога товаров (имитация)
    m_Products("ITEM001") = CreateObject("Scripting.Dictionary")
    m_Products("ITEM001")("Name") = "Супер-виджет"
    m_Products("ITEM001")("Price") = 10.50
    m_Products("ITEM001")("Stock") = 100
    
    m_Products("ITEM002") = CreateObject("Scripting.Dictionary")
    m_Products("ITEM002")("Name") = "Мега-гаджет"
    m_Products("ITEM002")("Price") = 75.25
    m_Products("ITEM002")("Stock") = 50

    m_Products("SOLD_OUT") = CreateObject("Scripting.Dictionary")
    m_Products("SOLD_OUT")("Name") = "Эксклюзив (Распродан)"
    m_Products("SOLD_OUT")("Price") = 99.99
    m_Products("SOLD_OUT")("Stock") = 0
    LoggerService.GetLogger.Log "CProductCatalog: Каталог товаров инициализирован. Загружено товаров: " & m_Products.Count
End Sub

Public Function GetProductInfo(itemCode As String) As Object ' Scripting.Dictionary
    If m_Products.Exists(itemCode) Then
        Set GetProductInfo = m_Products(itemCode)
    Else
        Set GetProductInfo = Nothing
    End If
End Function

Public Function CheckStock(itemCode As String, quantity As Long) As Boolean
    Dim productInfo As Object
    Set productInfo = GetProductInfo(itemCode)
    If Not productInfo Is Nothing Then
        CheckStock = (productInfo("Stock") >= quantity)
    Else
        CheckStock = False ' Товар не найден
    End If
End Function

Public Sub UpdateStock(itemCode As String, quantityChange As Long) ' quantityChange может быть отрицательным
    If m_Products.Exists(itemCode) Then
        m_Products(itemCode)("Stock") = m_Products(itemCode)("Stock") + quantityChange
        LoggerService.GetLogger.Log "CProductCatalog: Сток для " & itemCode & " изменен на " & quantityChange & ". Новый сток: " & m_Products(itemCode)("Stock")
    End If
End Sub

Private Sub Class_Terminate()
    Set m_Products = Nothing
    LoggerService.GetLogger.Log "CProductCatalog: Экземпляр уничтожен."
End Sub
````

**2. Обновленный `COrderProcessingFacade` (использует `CProductCatalog` вместо `CInventoryManager`)**
*   Предположим, `CInventoryManager` был упрощен или его функциональность перенесена/объединена с `CProductCatalog` для этого примера. Для простоты, заменим `m_Inventory` на `m_Catalog`.

````vba
' COrderProcessingFacade - Модуль класса (Обновленный)
Option Explicit
Private m_Catalog As CProductCatalog ' Замена CInventoryManager
Private m_Payment As CPaymentProcessor
Private m_Notifier As CNotificationService

Private Sub Class_Initialize()
    Set m_Catalog = New CProductCatalog
    Set m_Payment = New CPaymentProcessor
    Set m_Notifier = New CNotificationService
    LoggerService.GetLogger.Log "COrderProcessingFacade: Инициализирован с CProductCatalog."
End Sub

Public Function GetProductDetails(itemCode As String) As String
    Dim productInfo As Object
    Set productInfo = m_Catalog.GetProductInfo(itemCode)
    If Not productInfo Is Nothing Then
        GetProductDetails = productInfo("Name") & " - Цена: " & FormatCurrency(productInfo("Price")) & " - На складе: " & productInfo("Stock")
    Else
        GetProductDetails = "Товар с кодом '" & itemCode & "' не найден."
    End If
End Function

Public Function PlaceOrder(customerEmail As String, itemCode As String, quantity As Long, paymentDetails As String) As Boolean
    Dim success As Boolean
    success = False
    LoggerService.GetLogger.Log "Facade: Начало оформления заказа для " & customerEmail & ", товар: " & itemCode & ", кол-во: " & quantity
    
    Dim productInfo As Object
    Set productInfo = m_Catalog.GetProductInfo(itemCode)
    
    If productInfo Is Nothing Then
        LoggerService.GetLogger.Log "Facade: Товар " & itemCode & " не найден в каталоге.", "ERROR"
        m_Notifier.SendFailureNotice customerEmail, "Не удалось оформить заказ: товар " & itemCode & " не найден."
        PlaceOrder = False
        Exit Function
    End If
    
    Dim unitPrice As Double
    unitPrice = productInfo("Price")

    ' 1. Проверить наличие товара
    If Not m_Catalog.CheckStock(itemCode, quantity) Then
        LoggerService.GetLogger.Log "Facade: Товара " & itemCode & " недостаточно на складе (нужно " & quantity & ", есть " & productInfo("Stock") & ").", "WARN"
        m_Notifier.SendFailureNotice customerEmail, "Не удалось оформить заказ: товар " & productInfo("Name") & " отсутствует в нужном количестве."
        PlaceOrder = False
        Exit Function
    End If
    LoggerService.GetLogger.Log "Facade: Товар " & itemCode & " в наличии (" & productInfo("Stock") & " шт.)."

    ' 2. Зарезервировать товар (уменьшить сток)
    m_Catalog.UpdateStock itemCode, -quantity ' Уменьшаем сток
    LoggerService.GetLogger.Log "Facade: Сток для " & itemCode & " уменьшен на " & quantity

    ' 3. Обработать платеж
    If Not m_Payment.ProcessPayment(quantity * unitPrice, paymentDetails) Then
        LoggerService.GetLogger.Log "Facade: Платеж не прошел.", "ERROR"
        m_Catalog.UpdateStock itemCode, quantity ' Откат резервации (возвращаем сток)
        LoggerService.GetLogger.Log "Facade: Резерв товара " & itemCode & " отменен, сток восстановлен."
        m_Notifier.SendFailureNotice customerEmail, "Не удалось оформить заказ: проблема с оплатой."
        PlaceOrder = False
        Exit Function
    End If
    LoggerService.GetLogger.Log "Facade: Платеж (" & FormatCurrency(quantity * unitPrice) & ") успешно обработан."

    ' 4. Отправить подтверждение
    m_Notifier.SendConfirmation customerEmail, "Ваш заказ на " & productInfo("Name") & " (" & quantity & " шт.) по цене " & FormatCurrency(unitPrice) & " за ед. успешно оформлен и оплачен."
    LoggerService.GetLogger.Log "Facade: Подтверждение заказа отправлено на " & customerEmail, "SUCCESS"
    
    success = True
    PlaceOrder = success
End Function

Private Sub Class_Terminate()
    Set m_Catalog = Nothing
    Set m_Payment = Nothing
    Set m_Notifier = Nothing
    LoggerService.GetLogger.Log "COrderProcessingFacade: Деинициализирован."
End Sub

' CPaymentProcessor, CNotificationService, LoggerService, CAppLogger остаются как в предыдущих примерах
' Убедитесь, что LoggerService.GetLogger() доступен и логгер инициализирован (например, через AppBootstrapper)
````

**3. Клиентский код (например, в стандартном модуле `UserInterface`)**
````vba
' Module: UserInterface
Option Explicit

Sub ShowProductAndOrder()
    ' Предполагается, что AppBootstrapper.Execute был вызван и логгер инициализирован.
    ' Если нет, добавьте:
    ' If Not AppBootstrapper.IsBootstrapped() Then AppBootstrapper.Execute ThisWorkbook
    
    Dim facade As COrderProcessingFacade
    Set facade = New COrderProcessingFacade
    
    Dim itemCodeToQuery As String
    itemCodeToQuery = InputBox("Введите код товара для просмотра (например, ITEM001, ITEM002, SOLD_OUT):", "Запрос информации о товаре", "ITEM001")
    If itemCodeToQuery = "" Then
        Set facade = Nothing
        Exit Sub
    End If
    
    MsgBox facade.GetProductDetails(itemCodeToQuery), vbInformation, "Информация о товаре"
    
    If MsgBox("Хотите заказать этот товар?", vbYesNo + vbQuestion, "Оформление заказа") = vbYes Then
        Dim qty As Long
        Dim qtyStr As String
        qtyStr = InputBox("Введите количество:", "Количество товара", "1")
        If Not IsNumeric(qtyStr) Or CLng(qtyStr) <= 0 Then
            MsgBox "Некорректное количество.", vbExclamation
            Set facade = Nothing
            Exit Sub
        End If
        qty = CLng(qtyStr)
        
        Dim email As String
        email = InputBox("Введите ваш email:", "Email для подтверждения", "user@example.com")
        If email = "" Then
            Set facade = Nothing
            Exit Sub
        End If
        
        Dim paymentInfo As String ' В реальном приложении это были бы более сложные данные
        paymentInfo = "CARD_OK_1234" 
        
        If facade.PlaceOrder(email, itemCodeToQuery, qty, paymentInfo) Then
            MsgBox "Заказ успешно размещен!", vbInformation
        Else
            MsgBox "Не удалось разместить заказ. Подробности смотрите в логе.", vbExclamation
        End If
    End If
    
    Set facade = Nothing
End Sub
````

**Как это работает:**
1.  Клиентский код (`UserInterface.ShowProductAndOrder`) создает экземпляр `COrderProcessingFacade`.
2.  Клиент запрашивает информацию о товаре через `facade.GetProductDetails()`. Фасад делегирует этот запрос своему экземпляру `CProductCatalog`.
3.  Клиент размещает заказ через `facade.PlaceOrder()`. Фасад координирует работу `CProductCatalog` (проверка и обновление стока), `CPaymentProcessor` (обработка платежа) и `CNotificationService` (отправка уведомлений).
4.  Вся сложность взаимодействия между `CProductCatalog`, `CPaymentProcessor` и `CNotificationService` скрыта от `UserInterface`. Клиент работает только с простым и понятным интерфейсом Фасада.
5.  Все операции логируются через `LoggerService.GetLogger()`, который является Singleton'ом и был инициализирован ранее (например, в `AppBootstrapper`).

Этот мастер-класс показывает, как Фасад упрощает клиентский код, инкапсулируя сложную логику взаимодействия компонентов подсистемы и предоставляя единую точку доступа к ее функциональности.

---

## Заключение

Правильная инициализация и грамотная организация доступа к функциональности являются краеугольными камнями надежной архитектуры VBA-приложений. Понимание специфики VBA, такое как отсутствие традиционной точки входа `main()` и особенности управления объектами, позволяет избегать распространенных ошибок.

Применение таких паттернов, как **Bootstrapper** для централизованной инициализации, **Singleton** для управления глобальными ресурсами, **Lazy Initialization** для оптимизации производительности, а также четкое разделение ответственности между **стандартными модулями и модулями классов** и использование **Фасадов** для упрощения взаимодействия со сложными подсистемами, значительно повышает качество, поддерживаемость и масштабируемость VBA-кода.

ИИ-ассистентам следует уделять особое внимание этим аспектам, чтобы генерировать код, который не только работает, но и соответствует лучшим практикам разработки на VBA, адаптированным к его уникальной среде.

Следующий этап в построении надежной архитектуры – это управление жизненным циклом объектов и применение техник внедрения зависимостей, которые будут рассмотрены во второй части данного руководства.

---

## Список использованных источников (иллюстративный)

*Это иллюстративный список типов источников, которые были бы полезны при составлении такого руководства. Реальное составление потребовало бы глубокого анализа актуальных обсуждений и документации.*

1.  **Официальная документация Microsoft (MSDN/Docs):**
    *   Разделы по VBA для Office (Excel, Word, Access).
    *   Справочники по объектным моделям приложений Office.
    *   Статьи о COM и его взаимодействии с VBA.
    *   *Ключевые моменты:* Основы синтаксиса, объектная модель, события, управление ошибками.
2.  **Специализированные форумы и сообщества:**
    *   **StackOverflow (теги `vba`, `excel-vba`, etc.):**
        *   *Ключевые моменты:* Решения конкретных проблем, обсуждение паттернов, примеры кода, типичные ошибки и их исправления.
    *   **RubberduckVBA (GitHub Issues, Gitter Chat, документация):**
        *   *Ключевые моменты:* Инструменты для рефакторинга, статического анализа, юнит-тестирования в VBA; обсуждения лучших практик и чистого кода в VBA.
    *   **Форумы MrExcel, OzGrid, VBAExpress:**
        *   *Ключевые моменты:* Практические примеры, сложные сценарии, обмен опытом между разработчиками.
3.  **Книги по VBA и разработке ПО:**
    *   Walkenbach, John. "Excel VBA Programming For Dummies."
    *   McFedries, Paul. "VBA for Modelers: Developing Decision Support Systems with Microsoft Office Excel."
    *   Martin, Robert C. "Clean Code: A Handbook of Agile Software Craftsmanship" (принципы применимы с адаптацией).
    *   Gamma, E., Helm, R., Johnson, R., Vlissides, J. "Design Patterns: Elements of Reusable Object-Oriented Software" (концепции паттернов, адаптируемые для VBA).
    *   *Ключевые моменты:* Основы программирования, продвинутые техники, паттерны проектирования, принципы чистого кода.
4.  **Блоги и статьи опытных VBA-разработчиков:**
    *   Например, блоги экспертов, публикующих статьи о структурах данных в VBA, ООП в VBA, оптимизации производительности.
    *   *Ключевые моменты:* Нестандартные решения, глубокий анализ специфики VBA, практические советы.
5.  **Академические статьи (если применимо):**
    *   Редко напрямую затрагивают VBA, но могут касаться общих концепций управления жизненным циклом, DI, архитектуры ПО, которые можно адаптировать.
    *   *Ключевые моменты:* Теоретические основы, формальные подходы.

Приоритет при отборе источников отдавался бы обсуждениям архитектуры, концепций инициализации, управления жизненным циклом, инстанцирования классов, зависимостей в VBA и ООП-подходов, адаптированных для VBA, с акцентом на практическую применимость и избегание ошибок, часто совершаемых ИИ.


# Проектирование архитектуры приложений VBA: Управление жизненным циклом и паттерны внедрения зависимостей

## Содержание

4.  [Раздел 4: Управление жизненным циклом объектов](#раздел-4-управление-жизненным-циклом-объектов)
    *   [4.1. Подраздел: Создание и освобождение объектов (`Set New`, `Set Nothing`)](#41-подраздел-создание-и-освобождение-объектов-set-new-set-nothing)
    *   [4.2. Подраздел: Область видимости и время жизни переменных](#42-подраздел-область-видимости-и-время-жизни-переменных)
    *   [4.3. Подраздел: Избегание циклических ссылок](#43-подраздел-избегание-циклических-ссылок)
    *   [Резюме Раздела 4](#резюме-раздела-4)
    *   [Золотые правила Раздела 4](#золотые-правила-раздела-4)
    *   [Мастер-класс Раздела 4](#мастер-класс-раздела-4)
5.  [Раздел 5: Внедрение зависимостей (Dependency Injection) в VBA](#раздел-5-внедрение-зависимостей-dependency-injection-в-vba)
    *   [5.1. Подраздел: Что такое внедрение зависимостей и зачем оно в VBA?](#51-подраздел-что-такое-внедрение-зависимостей-и-зачем-оно-в-vba)
    *   [5.2. Подраздел: "Constructor Injection" (через процедуру Initialize/Configure)](#52-подраздел-constructor-injection-через-процедуру-initializeconfigure)
    *   [5.3. Подраздел: Property Injection (Внедрение через свойство)](#53-подраздел-property-injection-внедрение-через-свойство)
    *   [5.4. Подраздел: Parameter Injection (Внедрение через параметр метода)](#54-подраздел-parameter-injection-внедрение-через-параметр-метода)
    *   [5.5. Подраздел: Service Locator (Локатор служб)](#55-подраздел-service-locator-локатор-служб)
    *   [Резюме Раздела 5](#резюме-раздела-5)
    *   [Золотые правила Раздела 5](#золотые-правила-раздела-5)
    *   [Мастер-класс Раздела 5](#мастер-класс-раздела-5)
6.  [Раздел 6: Паттерны проектирования, адаптированные для VBA](#раздел-6-паттерны-проектирования-адаптированные-для-vba)
    *   [6.1. Подраздел: Factory (Фабрика объектов)](#61-подраздел-factory-фабрика-объектов)
    *   [6.2. Подраздел: Observer (Наблюдатель)](#62-подраздел-observer-наблюдатель)
    *   [Резюме Раздела 6](#резюме-раздела-6)
    *   [Золотые правила Раздела 6](#золотые-правила-раздела-6)
    *   [Мастер-класс Раздела 6](#мастер-класс-раздела-6)
7.  [Общее заключение](#общее-заключение)
8.  [Список использованных источников (иллюстративный)](#список-использованных-источников-иллюстративный-часть-2)

---

## Раздел 4: Управление жизненным циклом объектов

### 4.1. Подраздел: Создание и освобождение объектов (`Set New`, `Set Nothing`)

*   **Проблема:** Утечки памяти, "зависшие" или некорректно работающие объекты, ошибки "Object variable or With block variable not set" (Ошибка 91) при попытке доступа к объекту, который должен был быть освобожден, или наоборот, к объекту, который не был должным образом уничтожен и мешает работе.
*   **Причина (в контексте VBA):** VBA использует подсчет ссылок для COM-объектов (включая объекты приложений Office, такие как `Worksheet`, `Range`) и для экземпляров собственных классов. Когда счетчик ссылок на объект достигает нуля, объект уничтожается, и вызывается его метод `Class_Terminate` (если это экземпляр класса VBA). Неправильное управление ссылками (не освобождение объектов, когда они больше не нужны) приводит к тому, что счетчик ссылок не обнуляется, и объект остается в памяти.
*   **Ошибка (специфический симптом/пример):**
    *   Глобальная объектная переменная или переменная уровня модуля инициализируется, используется, но никогда не устанавливается в `Nothing`. Если это объект с ресурсами (например, открытый файл в `FileSystemObject`), ресурс может остаться заблокированным.
    *   Объекты добавляются в коллекцию, но сама коллекция или отдельные ее элементы не очищаются должным образом, удерживая ссылки.
    *   Приложение Excel не закрывается полностью или "подвисает" после выполнения макроса, так как в памяти остались активные COM-объекты с ненулевым счетчиком ссылок.
    ````vba
    ' Module1
    Public g_AppWideCollection As Collection ' Глобальная коллекция
    Public g_FSO As Object ' Глобальный FileSystemObject

    Sub InitializeGlobalObjects()
        Set g_AppWideCollection = New Collection
        Set g_FSO = CreateObject("Scripting.FileSystemObject")
        ' ... использование g_FSO ...
    End Sub

    Sub AddToCollection(item As Object)
        If g_AppWideCollection Is Nothing Then InitializeGlobalObjects
        g_AppWideCollection.Add item
    End Sub

    ' Проблема: Если не вызвать процедуру очистки, g_AppWideCollection и g_FSO
    ' (и все объекты в коллекции) останутся в памяти до полного сброса проекта VBA
    ' или закрытия Excel, потенциально удерживая ресурсы.
    ' g_FSO.OpenTextFile(...) - если файл не закрыт явно, он может остаться заблокированным.
    ````
*   **Решение (VBA-специфический паттерн/подход):**
    1.  **Явное создание:** Всегда используйте `Set obj = New ClassName` для экземпляров классов VBA или `Set obj = CreateObject("...")` / `Set obj = Application.Workbooks.Add` и т.д. для COM-объектов.
    2.  **Явное освобождение:** Когда объект больше не нужен, освободите ссылку на него, присвоив переменной значение `Nothing`: `Set obj = Nothing`. Это уменьшает счетчик ссылок на объект.
    3.  **`Class_Terminate`:** Используйте событие `Class_Terminate` в модулях классов для выполнения любой необходимой очистки непосредственно перед уничтожением объекта (например, закрытие файлов, освобождение других объектов, которые инкапсулированы данным классом).
    4.  **Очистка коллекций:** При работе с коллекциями, содержащими объекты, убедитесь, что либо сама коллекция устанавливается в `Nothing` (что освободит ссылки на все ее элементы, если это единственные ссылки), либо пройдитесь по коллекции и установите каждый элемент в `Nothing` перед очисткой самой коллекции, если это необходимо для немедленного вызова их `Class_Terminate`.
    5.  **Порядок освобождения:** Если объекты имеют зависимости, освобождайте их в порядке, обратном созданию или в соответствии с логикой зависимостей, чтобы избежать ошибок доступа к уже освобожденным зависимостям.
*   **Примеры кода:**

    **Плохо (Bad):** Забыли освободить объект.
    ````vba
    ' CMyResourceHolder - Модуль класса
    ' Private m_FileNum As Integer
    ' Private m_FilePath As String
    ' Public Sub OpenResource(filePath As String)
    '     m_FilePath = filePath
    '     m_FileNum = FreeFile
    '     Open m_FilePath For Output As #m_FileNum
    '     Print #m_FileNum, "Resource opened at " & Now()
    '     Debug.Print "CMyResourceHolder: Файл " & m_FilePath & " открыт."
    ' End Sub
    ' ' Отсутствует Class_Terminate или метод Close для закрытия файла
    
    ' Module1
    Sub UseResourceBadly()
        Dim holder As CMyResourceHolder
        Set holder = New CMyResourceHolder
        holder.OpenResource ThisWorkbook.Path & "\testfile.txt"
        ' ... какая-то работа ...
        ' Забыли: Set holder = Nothing
        ' Забыли: holder.CloseResource (если бы такой метод был)
        ' Файл testfile.txt может остаться открытым/заблокированным дольше необходимого.
        Debug.Print "UseResourceBadly: holder существует, файл может быть еще открыт."
    End Sub
    ````

    **Хорошо (Good):** Правильное создание, использование и освобождение.
    ````vba
    ' CMyResourceHolder - Модуль класса
    Option Explicit
    Private m_FileNum As Integer
    Private m_FilePath As String
    Private m_IsOpen As Boolean

    Public Sub OpenResource(filePath As String)
        If m_IsOpen Then CloseResource ' Закрыть предыдущий, если открыт
        
        m_FilePath = filePath
        On Error GoTo OpenErrorHandler
        m_FileNum = FreeFile
        Open m_FilePath For Output As #m_FileNum
        Print #m_FileNum, "Resource opened at " & Now()
        m_IsOpen = True
        Debug.Print "CMyResourceHolder: Файл " & m_FilePath & " открыт (FileNum: " & m_FileNum & ")."
        Exit Sub
    OpenErrorHandler:
        Debug.Print "CMyResourceHolder: Ошибка открытия файла " & m_FilePath & " - " & Err.Description
        m_IsOpen = False
    End Sub

    Public Sub WriteToResource(data As String)
        If Not m_IsOpen Then
            Debug.Print "CMyResourceHolder: Ресурс не открыт. Запись невозможна."
            Exit Sub
        End If
        Print #m_FileNum, data
        Debug.Print "CMyResourceHolder: Записано в файл: " & data
    End Sub

    Public Sub CloseResource()
        If m_IsOpen Then
            Close #m_FileNum
            m_IsOpen = False
            Debug.Print "CMyResourceHolder: Файл " & m_FilePath & " закрыт."
        End If
    End Sub

    Private Sub Class_Initialize()
        m_IsOpen = False
        Debug.Print "CMyResourceHolder: Экземпляр создан."
    End Sub

    Private Sub Class_Terminate()
        ' Гарантированное закрытие ресурса при уничтожении объекта
        CloseResource
        Debug.Print "CMyResourceHolder: Экземпляр уничтожен. Ресурс гарантированно закрыт."
    End Sub

    ' Module1
    Sub UseResourceProperly()
        Dim holder As CMyResourceHolder
        Set holder = New CMyResourceHolder
        
        holder.OpenResource ThisWorkbook.Path & "\testfile_good.txt"
        If Dir(ThisWorkbook.Path & "\testfile_good.txt") <> "" Then ' Проверка, что файл действительно создан
            holder.WriteToResource "Первая строка данных."
            holder.WriteToResource "Вторая строка данных."
        Else
             Debug.Print "UseResourceProperly: Файл не был создан, запись невозможна."
        End If
        
        ' Явное закрытие ресурса, когда он больше не нужен в текущей логике
        ' holder.CloseResource ' Можно вызвать, если ресурс нужно закрыть до уничтожения объекта holder

        ' Освобождение объекта. Class_Terminate будет вызван здесь (или когда последняя ссылка исчезнет).
        Set holder = Nothing
        Debug.Print "UseResourceProperly: holder установлен в Nothing."
    End Sub
    ````
*   **Резюме:** Тщательное управление созданием (`Set ... = New ...`) и освобождением (`Set ... = Nothing`) объектов является критически важным в VBA для предотвращения утечек памяти и ресурсов. Событие `Class_Terminate` предоставляет механизм для гарантированной очистки ресурсов, инкапсулированных классом.
*   **Контрольные правила:**
    *   `Must`: Всегда присваивать `Nothing` объектным переменным, когда они больше не нужны.
    *   `Must`: Реализовывать логику очистки (закрытие файлов, освобождение других объектов) в `Class_Terminate` для классов, управляющих внешними ресурсами или другими объектами.
    *   `Must Not`: Оставлять "висеть" ссылки на объекты, особенно глобальные или статические, без явного механизма их очистки.
    *   `Must`: Обращать внимание на порядок освобождения взаимозависимых объектов.

### 4.2. Подраздел: Область видимости и время жизни переменных

*   **Проблема:** Непредсказуемое поведение программы из-за неправильного понимания, когда переменные (особенно объектные) создаются, уничтожаются или сохраняют свои значения.
*   **Причина (в контексте VBA):** VBA имеет различные спецификаторы области видимости (`Dim`, `Private`, `Public`, `Friend`, `Global`) и времени жизни (`Static`).
    *   **Переменные уровня процедуры:** Объявленные с `Dim` существуют только во время выполнения процедуры. `Static` переменные уровня процедуры сохраняют свое значение между вызовами процедуры, но сбрасываются при сбросе проекта VBA (ошибка, стоп, изменение кода).
    *   **Переменные уровня модуля (стандартного или класса):** `Private` доступны только внутри модуля. `Public` (в стандартных модулях) доступны глобально. `Public` (в модулях классов) являются членами объекта. Их время жизни связано со временем жизни модуля (для стандартных) или экземпляра объекта (для классов). Глобальные переменные стандартных модулей сбрасываются при необработанной ошибке или сбросе проекта.
*   **Ошибка (специфический симптом/пример):**
    *   Использование `Static` объектной переменной в процедуре для реализации "ленивой" инициализации Singleton'а без учета того, что она может быть сброшена ошибкой, и при следующем вызове объект будет создан заново, нарушая принцип Singleton.
    *   Неожиданное сохранение состояния в `Static` переменной процедуры, приводящее к неверным результатам при повторных вызовах с другими входными данными.
    *   Попытка доступа к `Public` переменной стандартного модуля, ожидая, что она сохранила значение после ошибки, которая сбросила проект.
    ````vba
    ' Module1
    Function GetCounter() As Long
        Static s_Counter As Long ' Сохраняет значение между вызовами
        s_Counter = s_Counter + 1
        GetCounter = s_Counter
        ' Если произойдет необработанная ошибка где-либо в проекте,
        ' и проект сбросится, s_Counter вернется к 0.
    End Function

    Public g_AppStatus As String ' Глобальная переменная

    Sub TestScope()
        Dim localObj As Object
        Set localObj = New Collection ' Существует только внутри TestScope
        
        g_AppStatus = "Инициализировано"
        Debug.Print "Счетчик 1: " & GetCounter() ' 1
        Debug.Print "Счетчик 2: " & GetCounter() ' 2
        
        ' Имитация сброса проекта (например, нажатием кнопки "Стоп" в отладчике и затем "Продолжить"
        ' или возникновением необработанной ошибки в другом месте, а затем повторным вызовом TestScope).
        ' Если проект сброшен, то при следующем вызове GetCounter() вернет 1, а g_AppStatus будет пустой.
        ' Stop ' Раскомментируйте, запустите, нажмите F5. Затем закомментируйте Stop и запустите снова.
        
        Debug.Print "Статус приложения: " & g_AppStatus
        Set localObj = Nothing
    End Sub
    ````
*   **Решение (VBA-специфический паттерн/подход):**
    *   Четко выбирать область видимости: `Private` по умолчанию для членов класса и переменных модуля, если они не должны быть доступны извне. `Public` – для намеренно открытых интерфейсов.
    *   Использовать `Static` переменные в процедурах с осторожностью, понимая их поведение при сбросе проекта. Для хранения состояния, которое должно пережить отдельные вызовы процедур, но связано с конкретным экземпляром, использовать переменные-члены класса.
    *   Для состояния, которое должно быть действительно глобальным и устойчивым к некоторым сбросам (но не к закрытию приложения), предпочтительнее использовать переменные уровня модуля в стандартном модуле, но с четкой стратегией их инициализации и сброса (например, через Bootstrapper и Singleton-аксессор).
    *   Не полагаться на сохранение значений в `Public` переменных стандартных модулей после необработанных ошибок, если нет явной логики восстановления.
*   **Примеры кода:**

    **Плохо (Bad):** Ненадежный Singleton с использованием `Static` в функции.
    ````vba
    ' CMyService - Модуль класса
    ' Private m_CreationTime As Date
    ' Private Sub Class_Initialize()
    '     m_CreationTime = Now()
    '     Debug.Print "CMyService создан в: " & m_CreationTime
    ' End Sub
    ' Public Function GetCreationTime() As Date
    '     GetCreationTime = m_CreationTime
    ' End Function

    ' Module: ServiceLocator_Bad
    Function GetMyService_Static() As CMyService
        Static instance As CMyService
        If instance Is Nothing Then
            Debug.Print "GetMyService_Static: Создание нового экземпляра..."
            Set instance = New CMyService
        End If
        Set GetMyService_Static = instance
        ' Проблема: если проект сбрасывается (ошибка, Stop), 'instance' становится Nothing,
        ' и при следующем вызове создается НОВЫЙ экземпляр, нарушая Singleton.
    End Function

    ' Sub TestStaticSingletonIssue()
    '     Dim svc1 As CMyService
    '     Dim svc2 As CMyService
    '     Set svc1 = GetMyService_Static()
    '     Debug.Print "svc1 time: " & svc1.GetCreationTime()
    '
    '     ' Имитируем сброс проекта (нажмите Stop в IDE, затем F5 для продолжения)
    '     ' Stop
    '
    '     Set svc2 = GetMyService_Static() ' После сброса будет создан новый экземпляр
    '     Debug.Print "svc2 time: " & svc2.GetCreationTime()
    '     Debug.Print "svc1 Is svc2: " & (svc1 Is svc2) ' Будет False после сброса
    ' End Sub
    ````

    **Хорошо (Good):** Более надежный Singleton с переменной уровня модуля.
    ````vba
    ' CMyService - Модуль класса (тот же)
    ' Private m_CreationTime As Date
    ' Private Sub Class_Initialize()
    '     m_CreationTime = Now()
    '     Debug.Print "CMyService (Module Level Singleton) создан в: " & m_CreationTime
    ' End Sub
    ' Public Function GetCreationTime() As Date
    '     GetCreationTime = m_CreationTime
    ' End Function

    ' Module: ServiceLocator_Good
    Private m_ServiceInstance As CMyService ' Переменная уровня модуля

    Public Function GetMyService_ModuleLevel() As CMyService
        If m_ServiceInstance Is Nothing Then
            Debug.Print "GetMyService_ModuleLevel: Создание нового экземпляра..."
            Set m_ServiceInstance = New CMyService
        End If
        Set GetMyService_ModuleLevel = m_ServiceInstance
        ' Этот экземпляр также сбросится при необработанной ошибке, приводящей к сбросу проекта,
        ' но он более предсказуем, чем Static в процедуре для некоторых сценариев.
        ' Истинная устойчивость требует более сложных подходов или внешнего состояния.
    End Function
    
    Public Sub ResetMyService_ModuleLevel() ' Для тестов или явного сброса
        Debug.Print "GetMyService_ModuleLevel: Экземпляр сброшен."
        Set m_ServiceInstance = Nothing
    End Sub

    ' Sub TestModuleLevelSingleton()
    '     Dim svc1 As CMyService
    '     Dim svc2 As CMyService
    '
    '     ' Сначала сбросим, если предыдущий тест оставил экземпляр
    '     ' Call ServiceLocator_Good.ResetMyService_ModuleLevel
    '
    '     Set svc1 = ServiceLocator_Good.GetMyService_ModuleLevel()
    '     Debug.Print "svc1 time (Module): " & svc1.GetCreationTime()
    '
    '     ' Имитируем сброс проекта (нажмите Stop в IDE, затем F5 для продолжения)
    '     ' Stop
    '
    '     Set svc2 = ServiceLocator_Good.GetMyService_ModuleLevel()
    '     Debug.Print "svc2 time (Module): " & svc2.GetCreationTime()
    '     Debug.Print "svc1 Is svc2 (Module): " & (svc1 Is svc2)
    '     ' После сброса проекта и здесь будет False, так как m_ServiceInstance станет Nothing.
    '     ' Главное преимущество - явное управление через модуль, а не "магию" Static в процедуре.
    ' End Sub
    ````
*   **Резюме:** Понимание области видимости (`Dim`, `Private`, `Public`) и времени жизни (`Static`, уровень процедуры/модуля) переменных имеет решающее значение для написания предсказуемого и надежного VBA-кода. Неправильное использование может привести к трудно отлавливаемым ошибкам состояния и нарушению паттернов.
*   **Контрольные правила:**
    *   `Must`: Использовать наименьшую возможную область видимости для переменных.
    *   `Must`: Понимать, что `Static` переменные в процедурах и переменные уровня стандартного модуля сбрасываются при сбросе проекта VBA.
    *   `Must Not`: Полагаться на `Static` переменные в процедурах для реализации критически важных Singleton'ов, если не предусмотрена защита от сброса проекта.
    *   `Must`: Использовать переменные-члены класса для хранения состояния, специфичного для экземпляра объекта.

### 4.3. Подраздел: Избегание циклических ссылок

*   **Проблема:** Два или более объекта ссылаются друг на друга таким образом, что их счетчики ссылок никогда не достигают нуля, даже если внешние ссылки на них удалены (`Set obj = Nothing`). Это приводит к утечке памяти, так как объекты и их ресурсы не освобождаются, и их методы `Class_Terminate` не вызываются.
*   **Причина (в контексте VBA):** Если объект A хранит ссылку на объект B, а объект B хранит ссылку на объект A, то даже если все остальные ссылки на A и B будут удалены, эти два объекта будут "поддерживать жизнь" друг друга.
*   **Ошибка (специфический симптом/пример):**
    *   Класс `CParent` имеет член `Private m_Child As CChild`, а класс `CChild` имеет член `Private m_Parent As CParent`. При создании экземпляров `p As New CParent` и `c As New CChild`, и установке `p.SetChild c` и `c.SetParent p`, они образуют цикл. `Set p = Nothing` и `Set c = Nothing` не приведут к вызову их `Class_Terminate`.
    *   Приложение потребляет все больше памяти при многократном создании таких пар объектов.
    ````vba
    ' CParent - Модуль класса
    Option Explicit
    Private m_ChildObject As CChild
    Public Name As String
    
    Public Sub SetChild(childInstance As CChild)
        Set m_ChildObject = childInstance
        Debug.Print "CParent (" & Name & "): Ребенок установлен: " & childInstance.Name
    End Sub
    
    Private Sub Class_Initialize()
        Name = "Parent_" & Format(Timer, "0.00") ' Уникальное имя для отладки
        Debug.Print "CParent (" & Name & ") создан."
    End Sub
    
    Private Sub Class_Terminate()
        Debug.Print "CParent (" & Name & ") УНИЧТОЖЕН." ' Не будет вызван при циклической ссылке
        Set m_ChildObject = Nothing
    End Sub

    ' CChild - Модуль класса
    Option Explicit
    Private m_ParentObject As CParent
    Public Name As String
    
    Public Sub SetParent(parentInstance As CParent)
        Set m_ParentObject = parentInstance
        Debug.Print "CChild (" & Name & "): Родитель установлен: " & parentInstance.Name
    End Sub
    
    Private Sub Class_Initialize()
        Name = "Child_" & Format(Timer, "0.00") ' Уникальное имя для отладки
        Debug.Print "CChild (" & Name & ") создан."
    End Sub
    
    Private Sub Class_Terminate()
        Debug.Print "CChild (" & Name & ") УНИЧТОЖЕН." ' Не будет вызван при циклической ссылке
        Set m_ParentObject = Nothing
    End Sub

    ' Module: CyclicReferenceDemo
    Sub CreateCyclicReference()
        Dim p As CParent
        Dim c As CChild
        
        Debug.Print "--- Создание циклической ссылки ---"
        Set p = New CParent
        Set c = New CChild
        
        p.SetChild c ' p -> c
        c.SetParent p ' c -> p (цикл!)
        
        Debug.Print "--- Попытка освободить p и c ---"
        Set p = Nothing ' Уменьшает внешнюю ссылку на Parent, но Child все еще ссылается на него
        Set c = Nothing ' Уменьшает внешнюю ссылку на Child, но Parent все еще ссылается на него
        
        ' Class_Terminate для p и c НЕ БУДУТ вызваны. Объекты остаются в памяти.
        Debug.Print "--- CreateCyclicReference завершен. Проверьте окно Immediate на отсутствие сообщений Terminate. ---"
    End Sub
    ````
*   **Решение (VBA-специфический паттерн/подход):**
    1.  **Разрыв цикла вручную:** Перед освобождением объектов явно разорвать одну из ссылок в цикле. Например, добавить метод `Dispose` или `ClearReferences` в один или оба класса, который устанавливает внутреннюю ссылку на другой объект в `Nothing`. Этот метод должен быть вызван перед тем, как внешние ссылки на объекты будут установлены в `Nothing`.
    2.  **Слабые ссылки (Advanced, не нативно в VBA):** В некоторых языках есть концепция "слабых ссылок", которые не увеличивают счетчик ссылок. В VBA это можно имитировать, храня `ObjPtr(obj)` и восстанавливая объект по указателю, но это сложно, небезопасно и обычно не рекомендуется для типичных приложений.
    3.  **Архитектурное решение:** Пересмотреть дизайн. Возможно, дочернему объекту не нужна полная ссылка на родителя, а достаточно ID родителя или передача родителя как параметра метода, когда это необходимо. Использование событий для обратной связи от дочернего к родительскому объекту вместо прямой ссылки.
    4.  **Один владелец:** Четко определить, какой объект "владеет" другим и отвечает за его жизненный цикл. "Владеемый" объект не должен хранить сильную ссылку обратно на "владельца", если это создает цикл.

*   **Примеры кода:**

    **Плохо (Bad):** Циклическая ссылка без механизма разрыва (как в примере выше).

    **Хорошо (Good):** Добавление метода для разрыва цикла.
    ````vba
    ' CParentWithDispose - Модуль класса
    Option Explicit
    Private m_ChildObject As CChildWithDispose
    Public Name As String
    
    Public Sub SetChild(childInstance As CChildWithDispose)
        Set m_ChildObject = childInstance
        Debug.Print "CParentWithDispose (" & Name & "): Ребенок установлен: " & childInstance.Name
    End Sub
    
    Public Sub Dispose() ' Метод для разрыва цикла
        Debug.Print "CParentWithDispose (" & Name & "): Вызван Dispose. Обнуление ссылки на ребенка."
        Set m_ChildObject = Nothing
    End Sub
    
    Private Sub Class_Initialize()
        Name = "ParentD_" & Format(Timer, "0.00")
        Debug.Print "CParentWithDispose (" & Name & ") создан."
    End Sub
    
    Private Sub Class_Terminate()
        Debug.Print "CParentWithDispose (" & Name & ") УНИЧТОЖЕН."
        ' Убедимся, что ссылка на ребенка точно очищена, если Dispose не был вызван
        If Not m_ChildObject Is Nothing Then
            Debug.Print "CParentWithDispose (" & Name & ") Terminate: Дополнительная очистка ссылки на ребенка."
            ' Важно: если ребенок тоже имеет метод Dispose, его здесь вызывать опасно,
            ' т.к. это может привести к повторному входу или ошибкам, если ребенок уже уничтожается.
            ' Лучше, чтобы каждый объект сам управлял своими ссылками в своем Dispose/Terminate.
            Set m_ChildObject = Nothing
        End If
    End Sub

    ' CChildWithDispose - Модуль класса
    Option Explicit
    Private m_ParentObject As CParentWithDispose
    Public Name As String
    
    Public Sub SetParent(parentInstance As CParentWithDispose)
        Set m_ParentObject = parentInstance
        Debug.Print "CChildWithDispose (" & Name & "): Родитель установлен: " & parentInstance.Name
    End Sub
    
    Public Sub Dispose() ' Метод для разрыва цикла
        Debug.Print "CChildWithDispose (" & Name & "): Вызван Dispose. Обнуление ссылки на родителя."
        Set m_ParentObject = Nothing
    End Sub
    
    Private Sub Class_Initialize()
        Name = "ChildD_" & Format(Timer, "0.00")
        Debug.Print "CChildWithDispose (" & Name & ") создан."
    End Sub
    
    Private Sub Class_Terminate()
        Debug.Print "CChildWithDispose (" & Name & ") УНИЧТОЖЕН."
        If Not m_ParentObject Is Nothing Then
            Debug.Print "CChildWithDispose (" & Name & ") Terminate: Дополнительная очистка ссылки на родителя."
            Set m_ParentObject = Nothing
        End If
    End Sub

    ' Module: CyclicReferenceDemoGood
    Sub BreakCyclicReference()
        Dim p As CParentWithDispose
        Dim c As CChildWithDispose
        
        Debug.Print "--- Создание объектов с возможностью разрыва цикла ---"
        Set p = New CParentWithDispose
        Set c = New CChildWithDispose
        
        p.SetChild c
        c.SetParent p ' Цикл создан
        
        Debug.Print "--- Явный разрыв цикла перед освобождением ---"
        ' Разрываем цикл, вызвав Dispose у одного или обоих объектов.
        ' Порядок может быть важен в сложных сценариях.
        ' Обычно достаточно разорвать одну из ссылок.
        p.Dispose ' Parent больше не ссылается на Child
        ' c.Dispose ' Child больше не ссылается на Parent (можно и так, или оба)

        Debug.Print "--- Освобождение p и c ---"
        Set p = Nothing ' Теперь Parent может быть уничтожен, т.к. Child (если c.Dispose не вызван)
                      ' все еще может ссылаться на него, но внешняя ссылка и ссылка от Child (если p.Dispose вызван) удалены.
                      ' Если и p.Dispose, и c.Dispose были вызваны, то оба объекта не имеют взаимных ссылок.
        Set c = Nothing ' Теперь Child может быть уничтожен.
        
        ' Class_Terminate для p и c ТЕПЕРЬ БУДУТ вызваны.
        Debug.Print "--- BreakCyclicReference завершен. Проверьте окно Immediate на сообщения Terminate. ---"
    End Sub
    ````
*   **Резюме:** Циклические ссылки являются частой причиной утечек памяти в VBA. Их необходимо избегать на этапе проектирования или предусматривать явные механизмы для их разрыва (например, метод `Dispose`) перед освобождением объектов.
*   **Контрольные правила:**
    *   `Must`: Проектировать отношения между объектами так, чтобы избегать циклических сильных ссылок, если это возможно.
    *   `Must`: Если циклические ссылки неизбежны, реализовать метод (`Dispose`, `ClearReferences`, `Unlink`) для явного разрыва цикла перед тем, как переменные будут установлены в `Nothing`.
    *   `Must Not`: Полагаться на автоматическое освобождение объектов, если между ними существуют циклические ссылки.
    *   `Must`: В `Class_Terminate` освобождать только "дочерние" или агрегируемые объекты, но избегать вызова методов других объектов, которые могут быть в процессе уничтожения или уже уничтожены, особенно если это часть цикла.

### Резюме Раздела 4
Управление жизненным циклом объектов в VBA требует дисциплины. Явное создание (`Set New`) и освобождение (`Set Nothing`) объектов, понимание их области видимости и времени жизни, а также активное предотвращение или разрыв циклических ссылок – ключевые аспекты для создания стабильных и не допускающих утечек памяти приложений. Событие `Class_Terminate` является последним рубежом для очистки ресурсов, но оно не сработает, если объект удерживается циклической ссылкой.

### Золотые правила Раздела 4
1.  **Освобождай, что создал:** Всегда используй `Set obj = Nothing` для объектов, которые больше не нужны.
2.  **Знай свои границы:** Четко определяй область видимости и время жизни переменных, чтобы избежать неожиданного поведения.
3.  **Разрывай порочный круг:** Избегай циклических ссылок или предусматривай явные методы для их разрыва.
4.  **`Class_Terminate` – твой друг (но не панацея):** Используй `Class_Terminate` для финальной очистки, но помни, что он не вызовется при циклических ссылках.

### Мастер-класс Раздела 4: Система управления задачами с родительскими и дочерними объектами

**Сценарий:** Создадим простую систему, где есть `CTaskList` (родитель), который содержит коллекцию объектов `CTaskItem` (дети). Каждый `CTaskItem` должен иметь ссылку на свой `CTaskList` для некоторых операций (например, уведомление о завершении). Продемонстрируем создание, управление ссылками и правильное освобождение для избежания циклов.

**1. Модуль класса `ITaskEvents` (для обратной связи от Задачи к Списку)**
*   Это интерфейсный класс (без кода, только для `Implements`) для демонстрации одного из способов избежать прямой сильной ссылки от ребенка к родителю для обратных вызовов.

````vba
' ITaskEvents - Модуль класса (Интерфейс)
' Option Explicit
' Public Sub TaskCompleted(taskName As String)
' End Sub
' Public Sub TaskProgressUpdate(taskName As String, progress As Integer)
' End Sub
' В VBA интерфейсы создаются как модули классов, чьи публичные Sub/Function/Property
' затем реализуются (`Implements`) другими классами.
' Для простоты этого мастер-класса, мы можем не использовать интерфейс явно,
' а передавать родителя, но обеспечим разрыв ссылок.
' Однако, для более сложных сценариев, события или интерфейсы предпочтительнее.
````

**2. Модуль класса `CTaskItem`**
````vba
' CTaskItem - Модуль класса
Option Explicit

Private m_Name As String
Private m_IsCompleted As Boolean
Private m_ParentList As CTaskList ' Ссылка на родителя - потенциальный источник цикла

' Событие для уведомления (альтернатива прямой ссылке на родителя для вызова метода)
Public Event Completed(taskName As String)
Public Event ProgressUpdated(taskName As String, progressPercentage As Integer)

Public Property Get Name() As String
    Name = m_Name
End Property

Public Sub Initialize(taskName As String, parent As CTaskList)
    m_Name = taskName
    m_IsCompleted = False
    Set m_ParentList = parent ' Устанавливаем ссылку на родителя
    Debug.Print "CTaskItem '" & m_Name & "' создан и связан со списком '" & parent.ListName & "'."
End Sub

Public Sub MarkAsCompleted()
    If Not m_IsCompleted Then
        m_IsCompleted = True
        Debug.Print "CTaskItem '" & m_Name & "' помечен как выполненный."
        ' Уведомляем родителя или подписчиков
        RaiseEvent Completed(m_Name) ' Через событие
        ' Или, если бы была прямая зависимость и нужно было вызвать метод родителя:
        ' If Not m_ParentList Is Nothing Then
        ' m_ParentList.NotifyTaskCompleted Me ' Пример вызова метода родителя
        ' End If
    End If
End Sub

Public Sub UpdateProgress(percentage As Integer)
    Debug.Print "CTaskItem '" & m_Name & "' прогресс: " & percentage & "%"
    RaiseEvent ProgressUpdated(m_Name, percentage)
End Sub

Public Property Get IsCompleted() As Boolean
    IsCompleted = m_IsCompleted
End Property

Public Sub UnlinkParent() ' Метод для разрыва связи с родителем
    Debug.Print "CTaskItem '" & m_Name & "': разрыв связи с родителем."
    Set m_ParentList = Nothing
End Sub

Private Sub Class_Terminate()
    Debug.Print "CTaskItem '" & m_Name & "' УНИЧТОЖЕН."
    ' Убедимся, что ссылка на родителя очищена
    Set m_ParentList = Nothing
End Sub
````

**3. Модуль класса `CTaskList`**
````vba
' CTaskList - Модуль класса
Option Explicit
Private m_Tasks As Collection
Private m_ListName As String

' Для обработки событий от CTaskItem
Private WithEvents m_SampleTaskForEvents As CTaskItem ' Нужен экземпляр для привязки обработчиков

Public Property Get ListName() As String
    ListName = m_ListName
End Property

Private Sub Class_Initialize()
    Set m_Tasks = New Collection
    m_ListName = "СписокЗадач_" & CStr(Int(Timer * 100)) ' Уникальное имя
    Debug.Print "CTaskList '" & m_ListName & "' создан."
End Sub

Public Function AddTask(taskName As String) As CTaskItem
    Dim newTask As CTaskItem
    Set newTask = New CTaskItem
    newTask.Initialize taskName, Me ' Передаем себя как родителя -> создается ссылка Task -> List
                                    ' List хранит Task в коллекции m_Tasks -> создается ссылка List -> Task
                                    ' Это формирует цикл!
    m_Tasks.Add newTask, taskName
    
    ' Пример подключения к событиям ОДНОЙ задачи (для многих нужен массив WithEvents или класс-обертка)
    ' Этот механизм здесь больше для демонстрации, чем для полноценной обработки всех задач.
    If m_Tasks.Count = 1 Then ' Для простоты, подключимся к событиям первой добавленной задачи
        Set m_SampleTaskForEvents = newTask
    End If
    
    Set AddTask = newTask
End Function

Public Sub DisplayTasks()
    Debug.Print "--- Задачи в списке '" & m_ListName & "' (" & m_Tasks.Count & " шт.): ---"
    Dim task As CTaskItem
    If m_Tasks.Count = 0 Then
        Debug.Print "(пусто)"
    Else
        For Each task In m_Tasks
            Debug.Print " - " & task.Name & " (Выполнена: " & task.IsCompleted & ")"
        Next task
    End If
    Debug.Print "-------------------------------------"
End Sub

Public Sub ClearTasks() ' Метод для освобождения всех задач и разрыва циклов
    Debug.Print "CTaskList '" & m_ListName & "': Очистка задач..."
    Dim task As CTaskItem
    If Not m_Tasks Is Nothing Then
        For Each task In m_Tasks
            task.UnlinkParent ' Каждая задача разрывает свою ссылку на этот список
            Set task = Nothing  ' Освобождаем ссылку на задачу из коллекции (не обязательно, если коллекция потом Nothing)
        Next
        Set m_Tasks = New Collection ' Или Set m_Tasks = Nothing
        Debug.Print "CTaskList '" & m_ListName & "': Все задачи удалены и связи разорваны."
    End If
    Set m_SampleTaskForEvents = Nothing ' Отключаем обработчик событий
End Sub

' Обработчики событий от m_SampleTaskForEvents
Private Sub m_SampleTaskForEvents_Completed(taskName As String)
    Debug.Print "CTaskList '" & m_ListName & "' ПОЛУЧИЛ СОБЫТИЕ: Задача '" & taskName & "' выполнена!"
End Sub

Private Sub m_SampleTaskForEvents_ProgressUpdated(taskName As String, progressPercentage As Integer)
    Debug.Print "CTaskList '" & m_ListName & "' ПОЛУЧИЛ СОБЫТИЕ: Задача '" & taskName & "' прогресс: " & progressPercentage & "%"
End Sub

Private Sub Class_Terminate()
    Debug.Print "CTaskList '" & m_ListName & "': Попытка вызова Terminate..."
    ClearTasks ' Гарантированная очистка при уничтожении списка
    Set m_Tasks = Nothing
    Debug.Print "CTaskList '" & m_ListName & "' УНИЧТОЖЕН."
End Sub
````

**4. Стандартный модуль `TaskManagerDemo`**
````vba
' Module: TaskManagerDemo
Option Explicit

Sub RunTaskManagerDemo()
    Dim mainList As CTaskList
    Set mainList = New CTaskList
    
    Dim task1 As CTaskItem, task2 As CTaskItem, task3 As CTaskItem
    
    Set task1 = mainList.AddTask("Помыть посуду")
    Set task2 = mainList.AddTask("Вынести мусор")
    Set task3 = mainList.AddTask("Написать код для демо")
    
    mainList.DisplayTasks
    
    task1.UpdateProgress 50 ' Первая задача (m_SampleTaskForEvents) отправит событие
    task1.MarkAsCompleted    ' Первая задача (m_SampleTaskForEvents) отправит событие
    
    task2.MarkAsCompleted    ' Эта задача не m_SampleTaskForEvents, событие не будет обработано в CTaskList
                            ' но сама задача будет помечена как выполненная.
    
    mainList.DisplayTasks
    
    Debug.Print "--- Освобождение mainList ---"
    ' Если бы не было ClearTasks или task.UnlinkParent, здесь была бы утечка,
    ' так как mainList хранит task1/2/3, а task1/2/3 хранят mainList.
    
    ' mainList.ClearTasks ' Можно вызвать явно, если список должен быть очищен до его уничтожения.
                        ' Class_Terminate CTaskList также вызовет ClearTasks.
    
    Set mainList = Nothing ' Это инициирует Class_Terminate для mainList.
                         ' Внутри Terminate CTaskList вызовется ClearTasks.
                         ' ClearTasks вызовет UnlinkParent для каждой задачи.
                         ' После UnlinkParent и удаления из коллекции, задачи станут доступны для GC.

    ' Освобождаем прямые ссылки на задачи, если они еще есть (хотя в этом демо они уже не нужны после mainList = Nothing)
    Set task1 = Nothing
    Set task2 = Nothing
    Set task3 = Nothing
    
    Debug.Print "--- Демо завершено. Проверьте окно Immediate на сообщения Terminate. ---"
End Sub
````

**Как это работает:**
1.  `CTaskList` создает `CTaskItem` и передает ссылку на себя (`Me`) в метод `Initialize` задачи. Задача сохраняет эту ссылку.
2.  `CTaskList` добавляет задачу в свою коллекцию `m_Tasks`. Теперь `CTaskList` ссылается на `CTaskItem`, а `CTaskItem` ссылается на `CTaskList` – **цикл создан**.
3.  Чтобы разорвать цикл, у `CTaskItem` есть метод `UnlinkParent()`, который обнуляет его ссылку `m_ParentList`.
4.  У `CTaskList` есть метод `ClearTasks()`, который проходит по всем задачам и вызывает у каждой `task.UnlinkParent()`, а затем очищает свою коллекцию `m_Tasks`.
5.  Ключевой момент: `CTaskList.Class_Terminate` вызывает `ClearTasks()`. Это гарантирует, что при уничтожении списка задач (когда последняя внешняя ссылка на него исчезает, например, `Set mainList = Nothing`), все циклические ссылки с дочерними задачами будут разорваны, позволяя и списку, и задачам корректно уничтожиться.
6.  События (`Public Event Completed`) в `CTaskItem` и `WithEvents` в `CTaskList` показывают альтернативный способ обратной связи, который не создает сильных циклических ссылок (хотя `WithEvents` сама по себе создает ссылку от обработчика к источнику события).

Этот мастер-класс демонстрирует как возникновение циклических ссылок, так и один из способов их корректного разрешения через явные методы разрыва связей, обеспечивая правильное управление жизненным циклом объектов.

---

## Раздел 5: Внедрение зависимостей (Dependency Injection) в VBA

### 5.1. Подраздел: Что такое внедрение зависимостей и зачем оно в VBA?

*   **Проблема:** Классы тесно связаны друг с другом, так как они сами создают экземпляры своих зависимостей. Это затрудняет замену зависимостей (например, для тестирования или изменения поведения), увеличивает сложность кода и снижает его модульность.
*   **Причина (в контексте VBA):** Часто классы внутри своих методов или в `Class_Initialize` напрямую создают объекты, от которых они зависят, используя `Set dependency = New ConcreteDependencyClass`. Это называется жесткой связью (tight coupling).
*   **Ошибка (специфический симптом/пример):**
    *   Класс `CReportGenerator` всегда создает экземпляр `CExcelDataExporter` для экспорта данных. Если потребуется экспортировать данные в CSV или базу данных, придется изменять код `CReportGenerator` или создавать его наследников (что в VBA ограничено).
    *   При юнит-тестировании `CReportGenerator` невозможно подменить `CExcelDataExporter` на тестовый объект (mock/stub), который имитирует экспорт без реального взаимодействия с Excel.
    ````vba
    ' CDataService_HardcodedDependency - Модуль класса
    Option Explicit
    Private m_Logger As CFileLogger ' Жесткая зависимость

    Private Sub Class_Initialize()
        ' Класс сам создает свою зависимость
        Set m_Logger = New CFileLogger
        m_Logger.InitializeLog "CDataService.log"
        m_Logger.LogMessage "CDataService_HardcodedDependency: Инициализирован."
    End Sub

    Public Sub ProcessData(data As String)
        m_Logger.LogMessage "Processing data: " & data
        ' ... логика обработки данных ...
        m_Logger.LogMessage "Data processed."
    End Sub
    
    ' CFileLogger - Модуль класса (пример зависимости)
    ' Option Explicit
    ' Private FSO As Object
    ' Private LogStream As Object
    ' Public Sub InitializeLog(filePath As String)
    '     Set FSO = CreateObject("Scripting.FileSystemObject")
    '     Set LogStream = FSO.OpenTextFile(ThisWorkbook.Path & "\" & filePath, 8, True) ' 8 = Append, True = Create
    '     LogMessage "Logger initialized. Log file: " & filePath
    ' End Sub
    ' Public Sub LogMessage(message As String)
    '     If Not LogStream Is Nothing Then LogStream.WriteLine Now & " - " & message
    '     Debug.Print Now & " - " & message
    ' End Sub
    ' Private Sub Class_Terminate()
    '     If Not LogStream Is Nothing Then LogStream.Close
    '     Set FSO = Nothing
    ' End Sub
    ````
*   **Решение (VBA-специфический паттерн/подход):** Принцип инверсии зависимостей (Dependency Inversion Principle) – классы должны зависеть от абстракций (интерфейсов), а не от конкретных реализаций. Внедрение зависимостей (Dependency Injection, DI) – это техника, при которой зависимости объекта предоставляются ему извне, а не создаются им самим.
    *   **Преимущества DI:**
        *   **Слабая связанность (Loose Coupling):** Классы не зависят от конкретных реализаций своих зависимостей.
        *   **Улучшенная тестируемость:** Легко подменять реальные зависимости на тестовые двойники (mocks, stubs).
        *   **Большая гибкость и расширяемость:** Проще добавлять новые реализации зависимостей или изменять существующие, не затрагивая классы, которые их используют.
        *   **Повышение переиспользуемости компонентов.**
    *   **Основные типы DI (адаптированные для VBA):**
        1.  **"Constructor" Injection:** Зависимости передаются через специальный метод инициализации (например, `Initialize` или `Configure`), вызываемый сразу после создания объекта.
        2.  **Property Injection:** Зависимости устанавливаются через публичные свойства (`Property Set`).
        3.  **Parameter Injection:** Зависимости передаются как параметры методов, которые их используют.
*   **Резюме:** Внедрение зависимостей – это мощная техника для создания гибких, тестируемых и слабосвязанных VBA-приложений. Вместо того чтобы классы создавали свои зависимости, эти зависимости "внедряются" в них извне.
*   **Контрольные правила:**
    *   `Must`: Стремиться к тому, чтобы классы получали свои зависимости извне, а не создавали их сами.
    *   `Must`: Определять зависимости через абстракции (интерфейсы, если возможно, или базовые классы), а не конкретные классы, где это применимо. (В VBA "интерфейс" - это класс, реализуемый через `Implements`).
    *   `Must Not`: Создавать жесткие связи с конкретными классами зависимостей внутри методов или `Class_Initialize`, если эти зависимости могут меняться или требуют подмены для тестирования.

### 5.2. Подраздел: "Constructor Injection" (через процедуру Initialize/Configure)

*   **Проблема:** Как обеспечить объект всеми необходимыми зависимостями в момент его "создания" или непосредственно после, чтобы он был готов к работе.
*   **Причина (в контексте VBA):** В VBA нет настоящих конструкторов с параметрами, как в C# (`public MyClass(IDependency dep)`) или Java. Событие `Class_Initialize` выполняется без параметров.
*   **Ошибка (специфический симптом/пример):** Объект создается, но не может функционировать, пока его зависимости не будут установлены вручную через несколько отдельных вызовов методов или свойств, что делает объект временно невалидным. Или, что хуже, объект пытается работать без установленных зависимостей, приводя к ошибкам.
*   **Решение (VBA-специфический паттерн/подход):** Использовать публичный метод, условно называемый `Initialize`, `Configure`, `Setup` или `InjectDependencies`, который принимает все обязательные зависимости как параметры. Этот метод должен вызываться сразу после создания экземпляра объекта (`Set obj = New MyClass`).
*   **Примеры кода:**

    **Плохо (Bad):** Зависимости устанавливаются по частям, объект может быть не готов.
    ````vba
    ' CReportBuilder_Bad - Модуль класса
    ' Option Explicit
    ' Public DataSource As Object ' IDataSource
    ' Public Formatter As Object ' IReportFormatter
    ' Public Logger As Object    ' ILogger
    '
    ' ' Нет явного метода инициализации, зависимости нужно устанавливать по одной
    '
    ' Public Sub GenerateReport()
    '     If DataSource Is Nothing Or Formatter Is Nothing Or Logger Is Nothing Then
    '         Debug.Print "Ошибка: Не все зависимости установлены для CReportBuilder_Bad!"
    '         Exit Sub
    '     End If
    '     Logger.Log "Начало генерации отчета..."
    '     ' ... использует DataSource и Formatter ...
    '     Logger.Log "Отчет сгенерирован."
    ' End Sub
    '
    ' ' Клиентский код:
    ' ' Dim reportBuilder As CReportBuilder_Bad
    ' ' Set reportBuilder = New CReportBuilder_Bad
    ' ' ' Если забыть одну из следующих строк, GenerateReport не сработает или выдаст ошибку
    ' ' Set reportBuilder.DataSource = New CExcelDataSource
    ' ' Set reportBuilder.Formatter = New CSimpleTextFormatter
    ' ' Set reportBuilder.Logger = LoggerService.GetLogger ' Предположим, логгер - Singleton
    ' ' reportBuilder.GenerateReport
    ````

    **Хорошо (Good):** Использование метода `Initialize` для внедрения зависимостей.
    ````vba
    ' --- Абстракции/Интерфейсы (пустые классы для демонстрации) ---
    ' IDataSource - Модуль класса
    ' Public Function GetData() As Variant: End Function

    ' IReportFormatter - Модуль класса
    ' Public Function FormatData(data As Variant) As String: End Function

    ' ILogger - Модуль класса (может быть из предыдущих примеров, как CAppLogger)
    ' Public Sub Log(message As String, Optional logLevel As String = "INFO"): End Sub

    ' --- Конкретные реализации (заглушки) ---
    ' CExcelDataSource - Модуль класса
    ' Implements IDataSource
    ' Public Function IDataSource_GetData() As Variant: IDataSource_GetData = "Данные из Excel": End Function

    ' CSimpleTextFormatter - Модуль класса
    ' Implements IReportFormatter
    ' Public Function IReportFormatter_FormatData(data As Variant) As String: IReportFormatter_FormatData = "Отформатировано: " & CStr(data): End Function
    
    ' CConsoleLogger - Модуль класса
    ' Implements ILogger
    ' Public Sub ILogger_Log(message As String, Optional logLevel As String = "INFO"): Debug.Print Now & " [" & logLevel & "] " & message: End Sub


    ' CReportBuilder_Good - Модуль класса
    Option Explicit
    Private m_DataSource As IDataSource
    Private m_Formatter As IReportFormatter
    Private m_Logger As ILogger
    Private m_IsInitialized As Boolean

    ' "Конструктор" с внедрением зависимостей
    Public Sub Initialize(dataSourceDep As IDataSource, formatterDep As IReportFormatter, loggerDep As ILogger)
        If dataSourceDep Is Nothing Then Err.Raise 5, "CReportBuilder.Initialize", "DataSource не может быть Nothing"
        If formatterDep Is Nothing Then Err.Raise 5, "CReportBuilder.Initialize", "Formatter не может быть Nothing"
        If loggerDep Is Nothing Then Err.Raise 5, "CReportBuilder.Initialize", "Logger не может быть Nothing"
        
        Set m_DataSource = dataSourceDep
        Set m_Formatter = formatterDep
        Set m_Logger = loggerDep
        m_IsInitialized = True
        m_Logger.ILogger_Log "CReportBuilder_Good: Инициализирован со всеми зависимостями."
    End Sub

    Public Sub GenerateReport()
        If Not m_IsInitialized Then
            Debug.Print "CReportBuilder_Good: Ошибка - объект не инициализирован!"
            Err.Raise 9101, "CReportBuilder.GenerateReport", "Объект не был инициализирован. Вызовите Initialize."
            Exit Sub
        End If
        
        m_Logger.ILogger_Log "Начало генерации отчета..."
        Dim rawData As Variant
        rawData = m_DataSource.IDataSource_GetData()
        
        Dim formattedReport As String
        formattedReport = m_Formatter.IReportFormatter_FormatData(rawData)
        
        m_Logger.ILogger_Log "Отчет: " & formattedReport
        m_Logger.ILogger_Log "Отчет сгенерирован."
        ' В реальном приложении отчет бы куда-то выводился или сохранялся
    End Sub

    Private Sub Class_Initialize()
        m_IsInitialized = False ' Явно устанавливаем флаг
    End Sub

    Private Sub Class_Terminate()
        Set m_DataSource = Nothing
        Set m_Formatter = Nothing
        Set m_Logger = Nothing
        Debug.Print "CReportBuilder_Good: Уничтожен, зависимости освобождены."
    End Sub

    ' --- Клиентский код (например, в стандартном модуле) ---
    ' Sub TestReportBuilderConstructorInjection()
    '     Dim reportBuilder As CReportBuilder_Good
    '     Set reportBuilder = New CReportBuilder_Good
    '
    '     Dim excelSource As IDataSource
    '     Dim textFormatter As IReportFormatter
    '     Dim consoleLogger As ILogger
    '
    '     Set excelSource = New CExcelDataSource
    '     Set textFormatter = New CSimpleTextFormatter
    '     Set consoleLogger = New CConsoleLogger ' Или LoggerService.GetLogger, если он реализует ILogger
    '
    '     ' Внедряем зависимости сразу после создания
    '     reportBuilder.Initialize excelSource, textFormatter, consoleLogger
    '
    '     reportBuilder.GenerateReport
    '
    '     Set reportBuilder = Nothing
    '     Set excelSource = Nothing
    '     Set textFormatter = Nothing
    '     Set consoleLogger = Nothing
    ' End Sub
    ````
*   **Резюме:** "Constructor Injection" в VBA реализуется через публичный метод `Initialize` (или аналогичный), который принимает обязательные зависимости. Это гарантирует, что объект полностью сконфигурирован и готов к работе сразу после вызова этого метода.
*   **Контрольные правила:**
    *   `Must`: Использовать метод `Initialize` для внедрения обязательных зависимостей, необходимых для корректной работы объекта.
    *   `Must`: Вызывать метод `Initialize` сразу после создания экземпляра объекта (`Set obj = New ...`).
    *   `Must`: Проверять в методе `Initialize`, что переданные зависимости не `Nothing` (если они обязательны).
    *   `Must Not`: Позволять объекту существовать в невалидном состоянии из-за отсутствия обязательных зависимостей; если объект не может работать без них, его методы должны сигнализировать об ошибке или не выполняться до инициализации.

### 5.3. Подраздел: Property Injection (Внедрение через свойство)

*   **Проблема:** Как предоставить объекту опциональные зависимости или позволить изменять зависимости после первоначальной инициализации.
*   **Причина (в контексте VBA):** Не все зависимости являются строго обязательными для функционирования объекта, или может потребоваться их замена в процессе работы. "Constructor Injection" (через `Initialize`) обычно используется для обязательных зависимостей.
*   **Ошибка (специфический симптом/пример):**
    *   Метод `Initialize` становится слишком громоздким с большим количеством необязательных параметров.
    *   Отсутствует способ изменить зависимость (например, логгер) во время выполнения без пересоздания основного объекта.
*   **Решение (VBA-специфический паттерн/подход):** Использовать публичные свойства `Property Set` для установки зависимостей. Это позволяет клиенту устанавливать зависимости по мере необходимости.
*   **Примеры кода:**

    **Плохо (Bad):** Смешивание обязательных и множества опциональных зависимостей в `Initialize`.
    ````vba
    ' CTaskProcessor_BadProperties - Модуль класса
    ' Option Explicit
    ' Private m_PrimaryService As Object ' IPrimaryService (обязательный)
    ' Private m_OptionalLogger As Object ' ILogger (опциональный)
    ' Private m_CacheService As Object   ' ICache (опциональный)
    ' Private m_ErrorHandler As Object   ' IErrorHandler (опциональный)
    '
    ' ' Initialize становится перегруженным
    ' Public Sub Initialize(primarySvc As Object, _
    '                       Optional optLogger As Object = Nothing, _
    '                       Optional optCache As Object = Nothing, _
    '                       Optional optErrHandler As Object = Nothing)
    '     Set m_PrimaryService = primarySvc
    '     Set m_OptionalLogger = optLogger
    '     Set m_CacheService = optCache
    '     Set m_ErrorHandler = optErrHandler
    '     ' ...
    ' End Sub
    '
    ' Public Sub DoWork()
    '     If m_PrimaryService Is Nothing Then Exit Sub
    '     If Not m_OptionalLogger Is Nothing Then m_OptionalLogger.Log "Starting work..."
    '     ' ...
    ' End Sub
    ````

    **Хорошо (Good):** Обязательные зависимости через `Initialize`, опциональные – через `Property Set`.
    ````vba
    ' --- Абстракции/Интерфейсы (пустые классы для демонстрации) ---
    ' IPrimaryService - Модуль класса
    ' Public Sub Execute(): End Sub

    ' ILogger - Модуль класса (как в предыдущем примере)
    ' Public Sub Log(message As String, Optional logLevel As String = "INFO"): End Sub

    ' ICache - Модуль класса
    ' Public Sub Store(key As String, value As Variant): End Sub
    ' Public Function Retrieve(key As String) As Variant: End Function

    ' --- Конкретные реализации (заглушки) ---
    ' CRealPrimaryService - Модуль класса
    ' Implements IPrimaryService
    ' Public Sub IPrimaryService_Execute(): Debug.Print "CRealPrimaryService: Executed.": End Sub

    ' CConsoleLogger - Модуль класса (как в предыдущем примере)
    ' Implements ILogger
    ' Public Sub ILogger_Log(message As String, Optional logLevel As String = "INFO"): Debug.Print Now & " [" & logLevel & "] " & message: End Sub

    ' CMemoryCache - Модуль класса
    ' Implements ICache
    ' Private m_CacheDict As Object
    ' Private Sub Class_Initialize(): Set m_CacheDict = CreateObject("Scripting.Dictionary"): End Sub
    ' Public Sub ICache_Store(key As String, value As Variant): m_CacheDict(key) = value: Debug.Print "CMemoryCache: Stored " & key: End Sub
    ' Public Function ICache_Retrieve(key As String) As Variant: If m_CacheDict.Exists(key) Then ICache_Retrieve = m_CacheDict(key) Else ICache_Retrieve = Empty: Debug.Print "CMemoryCache: Retrieved " & key: End Function
    ' Private Sub Class_Terminate(): Set m_CacheDict = Nothing: End Sub


    ' CTaskProcessor_GoodProperties - Модуль класса
    Option Explicit
    Private m_PrimaryService As IPrimaryService ' Обязательная зависимость
    Private m_Logger As ILogger                 ' Опциональная зависимость
    Private m_Cache As ICache                   ' Опциональная зависимость
    Private m_IsInitialized As Boolean

    ' "Конструктор" для обязательных зависимостей
    Public Sub Initialize(primarySvcParam As IPrimaryService)
        If primarySvcParam Is Nothing Then Err.Raise vbObjectError + 513, "Initialize", "PrimaryService не может быть Nothing."
        Set m_PrimaryService = primarySvcParam
        m_IsInitialized = True
        LogInternal "CTaskProcessor: Инициализирован с PrimaryService."
    End Sub

    ' Свойство для внедрения опционального логгера
    Public Property Set Logger(value As ILogger)
        Set m_Logger = value
        LogInternal "CTaskProcessor: Logger установлен/изменен."
    End Property
    Public Property Get Logger() As ILogger
        Set Logger = m_Logger
    End Property

    ' Свойство для внедрения опционального кэша
    Public Property Set Cache(value As ICache)
        Set m_Cache = value
        LogInternal "CTaskProcessor: Cache установлен/изменен."
    End Property
    Public Property Get Cache() As ICache
        Set Cache = m_Cache
    End Property
    
    Public Sub DoWork(taskID As String)
        If Not m_IsInitialized Then
            Debug.Print "CTaskProcessor: Не инициализирован!"
            Exit Sub
        End If

        LogInternal "Начало выполнения задачи: " & taskID
        
        Dim cachedResult As Variant
        If Not m_Cache Is Nothing Then
            cachedResult = m_Cache.ICache_Retrieve(taskID)
        End If
        
        If Not IsEmpty(cachedResult) Then
            LogInternal "Результат для " & taskID & " найден в кэше: " & cachedResult
        Else
            m_PrimaryService.IPrimaryService_Execute
            LogInternal "PrimaryService выполнен для " & taskID
            If Not m_Cache Is Nothing Then
                m_Cache.ICache_Store taskID, "Результат_для_" & taskID
            End If
        End If
        LogInternal "Задача " & taskID & " завершена."
    End Sub

    Private Sub LogInternal(message As String) ' Внутренний метод для логирования, если логгер доступен
        If Not m_Logger Is Nothing Then
            m_Logger.ILogger_Log message
        Else
            Debug.Print "CTaskProcessor (no logger): " & message
        End If
    End Sub
    
    Private Sub Class_Initialize()
        m_IsInitialized = False
    End Sub

    Private Sub Class_Terminate()
        LogInternal "CTaskProcessor: Уничтожается."
        Set m_PrimaryService = Nothing
        Set m_Logger = Nothing
        Set m_Cache = Nothing
    End Sub

    ' --- Клиентский код ---
    ' Sub TestPropertyInjection()
    '     Dim processor As CTaskProcessor_GoodProperties
    '     Set processor = New CTaskProcessor_GoodProperties
    '
    '     Dim realService As IPrimaryService
    '     Set realService = New CRealPrimaryService
    '     processor.Initialize realService ' Внедрение обязательной зависимости
    '
    '     ' Внедрение опциональных зависимостей через свойства
    '     Dim consoleLogger As ILogger
    '     Set consoleLogger = New CConsoleLogger
    '     Set processor.Logger = consoleLogger ' Property Injection
    '
    '     Dim memoryCache As ICache
    '     Set memoryCache = New CMemoryCache
    '     Set processor.Cache = memoryCache   ' Property Injection
    '
    '     processor.DoWork "Task123"
    '     processor.DoWork "Task123" ' Второй раз должен взять из кэша
    '
    '     ' Можно изменить зависимость во время выполнения
    '     ' Set processor.Logger = Nothing ' Отключить логирование
    '     ' processor.DoWork "Task456"
    '
    '     Set processor = Nothing
    '     Set realService = Nothing
    '     Set consoleLogger = Nothing
    '     Set memoryCache = Nothing
    ' End Sub
    ````
*   **Резюме:** Property Injection подходит для установки опциональных зависимостей или для изменения зависимостей после инициализации объекта. Это делает конфигурацию объекта более гибкой.
*   **Контрольные правила:**
    *   `Must`: Использовать Property Injection для опциональных зависимостей.
    *   `Must`: Обеспечить, чтобы объект мог корректно работать (возможно, с ограниченной функциональностью или поведением по умолчанию), если опциональные зависимости не установлены.
    *   `Must Not`: Делать обязательные зависимости устанавливаемыми только через свойства, если это может привести к тому, что объект будет использоваться в невалидном состоянии. Для обязательных зависимостей предпочтительнее "Constructor Injection".

### 5.4. Подраздел: Parameter Injection (Внедрение через параметр метода)

*   **Проблема:** Зависимость требуется только для выполнения одного конкретного метода и не является постоянным состоянием объекта.
*   **Причина (в контексте VBA):** Делать такую временную зависимость членом класса (через Constructor или Property Injection) избыточно, так как она не нужна для других методов или для общего состояния объекта.
*   **Ошибка (специфический симптом/пример):** Класс хранит ссылку на сервис, который используется только в одном редком методе, увеличивая сложность управления жизненным циклом этого сервиса и самого класса.
*   **Решение (VBA-специфический паттерн/подход):** Передавать зависимость непосредственно как параметр того метода, который ее использует. Объект не хранит ссылку на эту зависимость постоянно.
*   **Примеры кода:**

    **Плохо (Bad):** Хранение временной зависимости как члена класса.
    ````vba
    ' CDataExporter_BadParam - Модуль класса
    ' Option Explicit
    ' Private m_ExportFormatter As Object ' IExportFormatter - используется только в ExportDataOnce
    '
    ' ' Форматтер устанавливается, даже если ExportDataOnce никогда не вызовется
    ' Public Property Set Formatter(value As Object)
    '     Set m_ExportFormatter = value
    ' End Property
    '
    ' Public Sub ExportDataOnce(data As Variant, destination As String)
    '     If m_ExportFormatter Is Nothing Then
    '         Debug.Print "Ошибка: Форматтер не установлен для ExportDataOnce!"
    '         Exit Sub
    '     End If
    '     Dim formattedData As String
    '     formattedData = m_ExportFormatter.Format(data) ' Предполагается, что IExportFormatter имеет метод Format
    '     Debug.Print "Экспорт в " & destination & ": " & formattedData
    ' End Sub
    '
    ' Public Sub AnotherOperation()
    '     Debug.Print "CDataExporter_BadParam: Выполняется другая операция, не требующая форматтера."
    ' End Sub
    ````

    **Хорошо (Good):** Передача зависимости через параметр метода.
    ````vba
    ' --- Абстракция/Интерфейс ---
    ' IExportFormatter - Модуль класса
    ' Public Function FormatForExport(dataToFormat As Variant) As String: End Function

    ' --- Конкретная реализация ---
    ' CSimpleXmlFormatter - Модуль класса
    ' Implements IExportFormatter
    ' Public Function IExportFormatter_FormatForExport(dataToFormat As Variant) As String
    '     IExportFormatter_FormatForExport = "<data>" & CStr(dataToFormat) & "</data>"
    ' End Function

    ' CDataExporter_GoodParam - Модуль класса
    Option Explicit

    ' Нет члена класса для форматтера

    Public Sub ExportDataWithFormatter(data As Variant, destination As String, formatter As IExportFormatter)
        If formatter Is Nothing Then
            Debug.Print "CDataExporter_GoodParam: Ошибка - форматтер не предоставлен для экспорта!"
            Err.Raise vbObjectError + 514, "ExportDataWithFormatter", "Formatter не может быть Nothing."
            Exit Sub
        End If
        
        Dim formattedData As String
        formattedData = formatter.IExportFormatter_FormatForExport(data)
        
        ' Имитация экспорта
        Debug.Print "CDataExporter_GoodParam: Экспорт в '" & destination & "'. Отформатированные данные: " & formattedData
        ' Например, запись в файл:
        ' Dim fso As Object, ts As Object
        ' Set fso = CreateObject("Scripting.FileSystemObject")
        ' Set ts = fso.CreateTextFile(destination, True)
        ' ts.Write formattedData
        ' ts.Close
    End Sub

    Public Sub PerformOtherTasks()
        Debug.Print "CDataExporter_GoodParam: Выполнение других задач, не связанных с форматированным экспортом."
    End Sub
    
    Private Sub Class_Initialize()
        Debug.Print "CDataExporter_GoodParam: Экземпляр создан."
    End Sub
    Private Sub Class_Terminate()
        Debug.Print "CDataExporter_GoodParam: Экземпляр уничтожен."
    End Sub

    ' --- Клиентский код ---
    ' Sub TestParameterInjection()
    '     Dim exporter As CDataExporter_GoodParam
    '     Set exporter = New CDataExporter_GoodParam
    '
    '     exporter.PerformOtherTasks
    '
    '     ' Создаем и передаем форматтер только когда он нужен
    '     Dim xmlFormatter As IExportFormatter
    '     Set xmlFormatter = New CSimpleXmlFormatter
    '
    '     Dim myData As String
    '     myData = "Это тестовые данные для экспорта"
    '
    '     exporter.ExportDataWithFormatter myData, ThisWorkbook.Path & "\exported_data.xml", xmlFormatter
    '     Debug.Print "Данные экспортированы. Проверьте файл exported_data.xml"
    '
    '     ' xmlFormatter может быть освобожден здесь, exporter не держит на него ссылку
    '     Set xmlFormatter = Nothing
    '     Set exporter = Nothing
    ' End Sub
    ````
*   **Резюме:** Parameter Injection является самым простым способом внедрения зависимостей, которые нужны только на время выполнения одного метода. Это уменьшает связанность и упрощает класс, так как ему не нужно хранить и управлять жизненным циклом таких временных зависимостей.
*   **Контрольные правила:**
    *   `Must`: Использовать Parameter Injection для зависимостей, которые не являются постоянным состоянием объекта и требуются только для конкретных операций.
    *   `Must`: Убедиться, что метод проверяет переданную зависимость на `Nothing`, если она обязательна для его работы.
    *   `Must Not`: Превращать все зависимости в параметры методов, если они являются неотъемлемой частью состояния объекта и используются многими его методами (в этом случае предпочтительнее Constructor или Property Injection).

### 5.5. Подраздел: Service Locator (Локатор служб)

*   **Проблема:** Как предоставить глобальный доступ к общим службам (например, логгер, сервис конфигурации) без необходимости "протаскивать" их через множество слоев объектов с помощью Constructor или Property Injection.
*   **Причина (в контексте VBA):** Иногда явное внедрение всех зависимостей может показаться громоздким, особенно для широко используемых сквозных служб. Service Locator предлагает централизованную точку для получения экземпляров служб.
*   **Ошибка (специфический симптом/пример):**
    *   Классы напрямую обращаются к глобальному объекту-локатору, чтобы получить свои зависимости. Это скрывает истинные зависимости класса (они не видны в его конструкторе или свойствах), что затрудняет понимание и тестирование.
    *   Service Locator может превратиться в "божественный объект" или свалку глобальных переменных, нарушая принципы инкапсуляции и слабой связанности.
    *   Тестирование класса, использующего Service Locator, становится сложным, так как локатор нужно настраивать или подменять глобально.
*   **Решение (VBA-специфический паттерн/подход):** Service Locator – это объект, который знает, как получить (создать или найти) нужную службу. Классы запрашивают у локатора необходимые им службы.
    *   **В VBA это часто реализуется через стандартный модуль с публичными функциями `GetServiceName() As IServiceName`.**
    *   **Предостережение:** Service Locator часто рассматривается как анти-паттерн, если используется чрезмерно, так как он приводит к неявным зависимостям и затрудняет тестирование. Его следует использовать с большой осторожностью, преимущественно для действительно глобальных, сквозных служб. Явное DI (Constructor, Property, Parameter Injection) обычно предпочтительнее.
*   **Примеры кода:**

    **Пример (показывающий механику, но с оговорками об использовании):**
    ````vba
    ' --- Абстракции/Интерфейсы ---
    ' IApplicationLogger - Модуль класса
    ' Public Sub LogEvent(eventMessage As String): End Sub

    ' IConfigurationService - Модуль класса
    ' Public Function GetSetting(settingName As String) As Variant: End Function

    ' --- Конкретные реализации (заглушки) ---
    ' CAppEventLogger - Модуль класса
    ' Implements IApplicationLogger
    ' Public Sub IApplicationLogger_LogEvent(eventMessage As String)
    '     Debug.Print "AppEventLogger: " & Now & " - " & eventMessage
    ' End Sub

    ' CIniFileConfiguration - Модуль класса
    ' Implements IConfigurationService
    ' Private m_Settings As Object ' Scripting.Dictionary
    ' Public Sub LoadConfig(filePath As String)
    '     Set m_Settings = CreateObject("Scripting.Dictionary")
    '     ' Имитация загрузки из INI
    '     m_Settings("AppName") = "Мое VBA Приложение"
    '     m_Settings("Version") = "1.0.SL"
    '     Debug.Print "CIniFileConfiguration: Загружен конфиг " & filePath
    ' End Sub
    ' Public Function IConfigurationService_GetSetting(settingName As String) As Variant
    '     If m_Settings.Exists(settingName) Then
    '         IConfigurationService_GetSetting = m_Settings(settingName)
    '     Else
    '         IConfigurationService_GetSetting = CVErr(xlErrNA)
    '     End If
    ' End Function


    ' GlobalServiceLocator - Стандартный модуль
    Option Explicit
    Private s_LoggerInstance As IApplicationLogger
    Private s_ConfigInstance As IConfigurationService

    ' Метод для конфигурации локатора (обычно вызывается из Bootstrapper)
    Public Sub ConfigureLocator(logger As IApplicationLogger, config As IConfigurationService)
        Set s_LoggerInstance = logger
        Set s_ConfigInstance = config
        s_LoggerInstance.IApplicationLogger_LogEvent "GlobalServiceLocator: Сконфигурирован."
    End Sub

    Public Function GetLogger() As IApplicationLogger
        If s_LoggerInstance Is Nothing Then
            ' Можно добавить логику создания по умолчанию или ошибку
            Debug.Print "GlobalServiceLocator: ВНИМАНИЕ! Logger не был сконфигурирован!"
            ' Set s_LoggerInstance = New CConsoleLogger ' Аварийный логгер
            ' Err.Raise vbObjectError + 515, "GetLogger", "Logger не сконфигурирован."
            ' Для примера просто вернем Nothing, что приведет к ошибке у клиента
        End If
        Set GetLogger = s_LoggerInstance
    End Function

    Public Function GetConfiguration() As IConfigurationService
        If s_ConfigInstance Is Nothing Then
            ' Debug.Print "GlobalServiceLocator: ВНИМАНИЕ! Configuration не был сконфигурирован!"
        End If
        Set GetConfiguration = s_ConfigInstance
    End Function
    
    Public Sub ResetLocator() ' Для тестов
        Set s_LoggerInstance = Nothing
        Set s_ConfigInstance = Nothing
        Debug.Print "GlobalServiceLocator: Сброшен."
    End Sub


    ' CSomeBusinessLogic - Модуль класса, использующий Service Locator
    Option Explicit
    Public Sub PerformAction()
        Dim logger As IApplicationLogger
        Set logger = GlobalServiceLocator.GetLogger() ' Получение зависимости через локатор
        
        Dim config As IConfigurationService
        Set config = GlobalServiceLocator.GetConfiguration()
        
        If logger Is Nothing Or config Is Nothing Then
            Debug.Print "CSomeBusinessLogic: Не удалось получить зависимости из локатора."
            Exit Sub
        End If

        logger.IApplicationLogger_LogEvent "CSomeBusinessLogic: PerformAction начато."
        
        Dim appName As String
        appName = config.IConfigurationService_GetSetting("AppName")
        
        logger.IApplicationLogger_LogEvent "Действие выполняется для приложения: " & appName
        ' ... другая логика ...
        logger.IApplicationLogger_LogEvent "CSomeBusinessLogic: PerformAction завершено."
    End Sub

    ' --- Bootstrapper или точка инициализации приложения ---
    ' Sub InitializeApplicationWithLocator()
    '     ' 1. Создаем конкретные экземпляры служб
    '     Dim appLogger As IApplicationLogger
    '     Set appLogger = New CAppEventLogger
    '
    '     Dim iniConfig As CIniFileConfiguration ' Реализует IConfigurationService
    '     Set iniConfig = New CIniFileConfiguration
    '     iniConfig.LoadConfig ThisWorkbook.Path & "\app.ini" ' Предполагается, что LoadConfig есть
    '
    '     ' 2. Конфигурируем локатор этими экземплярами
    '     GlobalServiceLocator.ConfigureLocator appLogger, iniConfig
    '
    '     ' 3. Теперь другие части приложения могут использовать локатор
    '     Dim logic As CSomeBusinessLogic
    '     Set logic = New CSomeBusinessLogic
    '     logic.PerformAction
    '
    '     ' Очистка (в реальном приложении - при закрытии)
    '     ' GlobalServiceLocator.ResetLocator
    '     ' Set appLogger = Nothing
    '     ' Set iniConfig = Nothing
    '     ' Set logic = Nothing
    ' End Sub
    ````
*   **Резюме:** Service Locator предоставляет централизованный доступ к службам, но его следует использовать с осторожностью, так как он может скрывать зависимости и усложнять тестирование. Явное внедрение зависимостей часто является лучшей альтернативой.
*   **Контрольные правила:**
    *   `Should`: Рассматривать Service Locator только для небольшого числа действительно глобальных, сквозных служб, если явное DI становится чрезмерно сложным.
    *   `Must`: Конфигурировать Service Locator на самом раннем этапе работы приложения (например, в Bootstrapper).
    *   `Must Not`: Превращать Service Locator в место для получения всех подряд зависимостей; это нарушает инкапсуляцию и слабую связанность.
    *   `Must`: Понимать, что использование Service Locator затрудняет юнит-тестирование классов, так как требует глобальной настройки или подмены локатора.

### Резюме Раздела 5
Внедрение зависимостей (DI) – ключевая практика для создания слабосвязанных, гибких и тестируемых VBA-приложений. Выбор между "Constructor" Injection (через метод `Initialize`), Property Injection и Parameter Injection зависит от того, являются ли зависимости обязательными, опциональными или временными. Service Locator может быть использован ограниченно для глобальных служб, но с пониманием его недостатков. Применение DI способствует лучшей организации кода и упрощает его поддержку и развитие.

### Золотые правила Раздела 5
1.  **Инвертируй зависимости:** Классы должны получать зависимости извне, а не создавать их сами.
2.  **Обязательное – через `Initialize`:** Используй "Constructor Injection" для зависимостей, без которых объект не может функционировать.
3.  **Опциональное – через `Property Set`:** Используй Property Injection для необязательных зависимостей или тех, что могут меняться.
4.  **Временное – через параметры метода:** Используй Parameter Injection для зависимостей, нужных только для одной операции.
5.  **Service Locator – с осторожностью:** Используй Service Locator умеренно, осознавая риски скрытых зависимостей и усложнения тестирования.

### Мастер-класс Раздела 5: Система обработки данных с различными типами DI

**Сценарий:** Создадим систему, где `CDataProcessor` (обработчик данных) имеет обязательную зависимость `IDataSource` и опциональную `ILogger`. Для одной из операций он также будет использовать `IOutputFormatter`, передаваемый через параметр метода.

**1. Абстракции (Интерфейсы - модули классов)**
````vba
' IDataSource - Модуль класса
Option Explicit
Public Function FetchData(sourceIdentifier As String) As Variant
End Function

' ILogger - Модуль класса
Option Explicit
Public Sub Log(message As String)
End Sub

' IOutputFormatter - Модуль класса
Option Explicit
Public Function Format(data As Variant) As String
End Function
````

**2. Конкретные реализации**
````vba
' CSimpleDataSource - Модуль класса
Option Explicit
Implements IDataSource

Private Function IDataSource_FetchData(sourceIdentifier As String) As Variant
    Debug.Print "CSimpleDataSource: Загрузка данных из '" & sourceIdentifier & "'"
    IDataSource_FetchData = "Это данные из " & sourceIdentifier
End Function

' CConsoleLogger - Модуль класса
Option Explicit
Implements ILogger

Private Sub ILogger_Log(message As String)
    Debug.Print Now & " [LOG]: " & message
End Sub

' CPlainTextFormatter - Модуль класса
Option Explicit
Implements IOutputFormatter

Private Function IOutputFormatter_Format(data As Variant) As String
    IOutputFormatter_Format = "Formatted Text: " & CStr(data)
End Function

' CNullLogger - Модуль класса (для случаев, когда логгер не нужен)
Option Explicit
Implements ILogger
Private Sub ILogger_Log(message As String)
    ' Ничего не делает
End Sub
````

**3. `CDataProcessor` (использует разные типы DI)**
````vba
' CDataProcessor - Модуль класса
Option Explicit

Private m_DataSource As IDataSource ' Обязательная, через Initialize
Private m_Logger As ILogger       ' Опциональная, через Property Set
Private m_IsInitialized As Boolean

' "Constructor" Injection для обязательной зависимости
Public Sub Initialize(dataSource As IDataSource)
    If dataSource Is Nothing Then Err.Raise vbObjectError + 1001, "Initialize", "DataSource не может быть Nothing."
    Set m_DataSource = dataSource
    
    ' Устанавливаем логгер по умолчанию, если он не будет внедрен позже
    Set m_Logger = New CNullLogger ' Безопасный логгер по умолчанию
    
    m_IsInitialized = True
    LogMessage "CDataProcessor: Инициализирован с DataSource. Используется NullLogger по умолчанию."
End Sub

' Property Injection для опциональной зависимости Logger
Public Property Set Logger(value As ILogger)
    If value Is Nothing Then
        Set m_Logger = New CNullLogger ' Если передали Nothing, ставим NullLogger
        LogMessage "CDataProcessor: Logger установлен в NullLogger."
    Else
        Set m_Logger = value
        LogMessage "CDataProcessor: Logger установлен." ' Это сообщение будет залогировано новым логгером
    End If
End Property
Public Property Get Logger() As ILogger
    Set Logger = m_Logger
End Property

' Метод, использующий Parameter Injection для форматтера
Public Sub ProcessAndOutputData(sourceID As String, outputFormatter As IOutputFormatter)
    If Not m_IsInitialized Then
        Debug.Print "CDataProcessor: Не инициализирован!"
        Exit Sub
    End If
    If outputFormatter Is Nothing Then
        LogMessage "Ошибка: outputFormatter не предоставлен для ProcessAndOutputData."
        Exit Sub
    End If

    LogMessage "Начало обработки и вывода данных для источника: " & sourceID
    Dim rawData As Variant
    rawData = m_DataSource.IDataSource_FetchData(sourceID)
    LogMessage "Данные получены: " & CStr(rawData)
    
    Dim formattedOutput As String
    formattedOutput = outputFormatter.IOutputFormatter_Format(rawData)
    LogMessage "Данные отформатированы: " & formattedOutput
    
    ' Имитация вывода
    MsgBox formattedOutput, vbInformation, "Результат обработки"
    LogMessage "Данные выведены."
End Sub

Public Sub SimpleProcess(sourceID As String)
    If Not m_IsInitialized Then
        Debug.Print "CDataProcessor: Не инициализирован!"
        Exit Sub
    End If
    LogMessage "Начало простой обработки для источника: " & sourceID
    Dim rawData As Variant
    rawData = m_DataSource.IDataSource_FetchData(sourceID)
    LogMessage "Простая обработка завершена. Получены данные: " & CStr(rawData)
End Sub

Private Sub LogMessage(message As String)
    ' m_Logger гарантированно не Nothing после Initialize (там ставится CNullLogger)
    m_Logger.ILogger_Log message
End Sub

Private Sub Class_Initialize()
    m_IsInitialized = False
    ' Не устанавливаем m_Logger здесь, чтобы продемонстрировать NullLogger по умолчанию
End Sub

Private Sub Class_Terminate()
    If Not m_Logger Is Nothing Then LogMessage "CDataProcessor: Уничтожается."
    Set m_DataSource = Nothing
    Set m_Logger = Nothing
End Sub
````

**4. Стандартный модуль `DI_DemoRunner`**
````vba
' Module: DI_DemoRunner
Option Explicit

Sub RunDataProcessorDemo()
    ' 1. Создание зависимостей
    Dim dataSource As IDataSource
    Set dataSource = New CSimpleDataSource
    
    Dim consoleLogger As ILogger
    Set consoleLogger = New CConsoleLogger
    
    Dim textFormatter As IOutputFormatter
    Set textFormatter = New CPlainTextFormatter
    
    ' 2. Создание основного объекта и "Constructor Injection"
    Dim processor As CDataProcessor
    Set processor = New CDataProcessor
    processor.Initialize dataSource ' Внедряем обязательный dataSource
                                    ' На этом этапе processor.Logger будет CNullLogger
    
    processor.SimpleProcess "Источник_А_без_полного_логирования" ' Будет использовать NullLogger

    ' 3. Property Injection для опционального логгера
    Set processor.Logger = consoleLogger ' Теперь используется CConsoleLogger
    processor.Logger.ILogger_Log "ConsoleLogger активирован для processor." ' Проверка
    
    ' 4. Вызов метода с Parameter Injection
    processor.ProcessAndOutputData "Источник_Б", textFormatter
    
    ' 5. Демонстрация изменения логгера обратно на NullLogger
    Set processor.Logger = Nothing ' Это установит внутренний m_Logger в CNullLogger
    processor.SimpleProcess "Источник_В_снова_без_логирования"
    
    ' 6. Очистка
    Set processor = Nothing
    Set dataSource = Nothing
    Set consoleLogger = Nothing
    Set textFormatter = Nothing
    
    Debug.Print "--- Демо DI завершено ---"
End Sub
````

**Как это работает:**
1.  `CDataProcessor` получает обязательную зависимость `IDataSource` через метод `Initialize` ("Constructor Injection"). Внутри `Initialize` также устанавливается `CNullLogger` по умолчанию.
2.  Опциональная зависимость `ILogger` может быть установлена позже через свойство `processor.Logger = consoleLogger` (Property Injection). Если передать `Nothing`, установится `CNullLogger`.
3.  Метод `ProcessAndOutputData` требует `IOutputFormatter`, который передается как параметр (Parameter Injection) и используется только внутри этого метода.
4.  Это демонстрирует гибкость DI: обязательные зависимости гарантированы при инициализации, опциональные могут быть добавлены или изменены, а временные передаются по мере необходимости.

Этот мастер-класс показывает, как различные техники DI могут сосуществовать и использоваться для создания гибкой и хорошо структурированной системы.

---
## Раздел 6: Паттерны проектирования, адаптированные для VBA

Некоторые классические паттерны проектирования могут быть успешно адаптированы для VBA, помогая решать типичные проблемы и улучшая архитектуру приложений. Рассмотрим несколько, особенно актуальных в контексте инициализации, управления объектами и зависимостями.

### 6.1. Подраздел: Factory (Фабрика объектов)

*   **Проблема:** Логика создания объектов становится сложной или разбросанной по коду. Клиентский код вынужден знать о конкретных классах-реализациях, что увеличивает связанность. Необходимо инкапсулировать процесс создания объектов, особенно если он зависит от каких-либо условий или конфигурации.
*   **Причина (в контексте VBA):** Нужно создавать экземпляры различных классов, реализующих общий интерфейс (или имеющих общую базовую функциональность), в зависимости от внешних факторов (например, пользовательский выбор, данные из файла конфигурации). Прямое использование `New ClassA`, `New ClassB` в клиентском коде делает его негибким.
*   **Ошибка (специфический симптом/пример):** Множество блоков `If/ElseIf` или `Select Case` в клиентском коде, отвечающих за создание разных типов объектов. При добавлении нового типа объекта приходится изменять все эти блоки.
    ````vba
    ' --- Интерфейс и классы ---
    ' IExporter - Модуль класса
    ' Public Sub Export(data As String, path As String): End Sub

    ' CExcelExporter - Модуль класса
    ' Implements IExporter
    ' Public Sub IExporter_Export(data As String, path As String): Debug.Print "Экспорт в Excel: " & data & " по пути " & path: End Sub

    ' CCsvExporter - Модуль класса
    ' Implements IExporter
    ' Public Sub IExporter_Export(data As String, path As String): Debug.Print "Экспорт в CSV: " & data & " по пути " & path: End Sub

    ' --- Клиентский код (Плохо) ---
    ' Sub ClientCode_WithoutFactory(exportType As String, dataToExport As String, filePath As String)
    '     Dim exporter As IExporter
    '
    '     If exportType = "Excel" Then
    '         Set exporter = New CExcelExporter
    '     ElseIf exportType = "CSV" Then
    '         Set exporter = New CCsvExporter
    '     Else
    '         Debug.Print "Неизвестный тип экспортера!"
    '         Exit Sub
    '     End If
    '
    '     exporter.IExporter_Export dataToExport, filePath
    '     Set exporter = Nothing
    ' End Sub
    ````
*   **Решение (VBA-специфический паттерн/подход):** Создать класс-Фабрику (`CExporterFactory`), который инкапсулирует логику выбора и создания конкретного экземпляра экспортера. Клиентский код обращается к фабрике, запрашивая объект по некоторому идентификатору (например, строке типа), и получает готовый экземпляр, реализующий общий интерфейс.
*   **Примеры кода:**

    **Плохо (Bad):** Логика создания размазана (как в примере выше).

    **Хорошо (Good):** Использование Фабрики.
    ````vba
    ' --- Интерфейс и классы (те же IExporter, CExcelExporter, CCsvExporter) ---
    ' IExporter - Модуль класса
    ' Option Explicit
    ' Public Sub Export(data As String, path As String)
    ' End Sub

    ' CExcelExporter - Модуль класса
    ' Option Explicit
    ' Implements IExporter
    ' Private Sub IExporter_Export(data As String, path As String)
    '     Debug.Print "CExcelExporter: Экспорт '" & data & "' в Excel файл '" & path & "'"
    '     ' Реальная логика экспорта в Excel
    ' End Sub
    ' Private Sub Class_Initialize(): Debug.Print "CExcelExporter: создан": End Sub
    ' Private Sub Class_Terminate(): Debug.Print "CExcelExporter: уничтожен": End Sub

    ' CCsvExporter - Модуль класса
    ' Option Explicit
    ' Implements IExporter
    ' Private Sub IExporter_Export(data As String, path As String)
    '     Debug.Print "CCsvExporter: Экспорт '" & data & "' в CSV файл '" & path & "'"
    '     ' Реальная логика экспорта в CSV
    ' End Sub
    ' Private Sub Class_Initialize(): Debug.Print "CCsvExporter: создан": End Sub
    ' Private Sub Class_Terminate(): Debug.Print "CCsvExporter: уничтожен": End Sub
    
    ' CXmlExporter - Модуль класса (новый тип для демонстрации расширяемости)
    ' Option Explicit
    ' Implements IExporter
    ' Private Sub IExporter_Export(data As String, path As String)
    '     Debug.Print "CXmlExporter: Экспорт '" & data & "' в XML файл '" & path & "'"
    ' End Sub
    ' Private Sub Class_Initialize(): Debug.Print "CXmlExporter: создан": End Sub
    ' Private Sub Class_Terminate(): Debug.Print "CXmlExporter: уничтожен": End Sub


    ' CExporterFactory - Модуль класса (Фабрика)
    Option Explicit

    Public Function CreateExporter(exportType As String) As IExporter
        Dim newExporter As IExporter
        
        Select Case LCase(exportType)
            Case "excel"
                Set newExporter = New CExcelExporter
            Case "csv"
                Set newExporter = New CCsvExporter
            Case "xml"  ' Легко добавить новый тип
                Set newExporter = New CXmlExporter
            Case Else
                Debug.Print "CExporterFactory: Неизвестный тип экспортера '" & exportType & "'. Возвращено Nothing."
                Set newExporter = Nothing ' Или можно возбуждать ошибку: Err.Raise vbObjectError + 100, , "Unknown exporter type"
        End Select
        
        Set CreateExporter = newExporter
    End Function
    Private Sub Class_Initialize(): Debug.Print "CExporterFactory: создана": End Sub
    Private Sub Class_Terminate(): Debug.Print "CExporterFactory: уничтожена": End Sub

    ' --- Клиентский код (Хорошо) ---
    ' Sub ClientCode_WithFactory(exportType As String, dataToExport As String, filePath As String)
    '     Dim factory As CExporterFactory
    '     Set factory = New CExporterFactory
    '
    '     Dim exporter As IExporter
    '     Set exporter = factory.CreateExporter(exportType) ' Клиент не знает о CExcelExporter или CCsvExporter
    '
    '     If Not exporter Is Nothing Then
    '         exporter.IExporter_Export dataToExport, filePath
    '         Set exporter = Nothing
    '     Else
    '         MsgBox "Не удалось создать экспортер для типа: " & exportType, vbExclamation
    '     End If
    '
    '     Set factory = Nothing
    ' End Sub
    '
    ' Sub TestFactory()
    '    Call ClientCode_WithFactory("Excel", "Данные для Excel", "C:\temp\report.xlsx")
    '    Call ClientCode_WithFactory("CSV", "Данные для CSV", "C:\temp\report.csv")
    '    Call ClientCode_WithFactory("XML", "Данные для XML", "C:\temp\report.xml")
    '    Call ClientCode_WithFactory("PDF", "Данные для PDF", "C:\temp\report.pdf") ' Неизвестный тип
    ' End Sub
    ````
*   **Резюме:** Паттерн Фабрика инкапсулирует логику создания объектов, делая систему более гибкой и уменьшая связанность клиентского кода с конкретными классами. Это упрощает добавление новых типов продуктов без изменения клиентского кода.
*   **Контрольные правила:**
    *   `Must`: Использовать Фабрику, когда логика создания объектов сложна, зависит от условий, или когда нужно скрыть от клиента конкретные классы-реализации.
    ' *   `Must`: Фабричный метод должен возвращать объект через его абстракцию (интерфейс или общий базовый класс).
    *   `Must Not`: Помещать в Фабрику бизнес-логику, не связанную с созданием объектов. Фабрика должна только создавать.

### 6.2. Подраздел: Observer (Наблюдатель)

*   **Проблема:** Необходимо, чтобы одни объекты (Наблюдатели, Subscribers) автоматически получали уведомления и обновлялись при изменении состояния другого объекта (Субъекта, Publisher), без того чтобы Субъект жестко зависел от конкретных классов Наблюдателей.
*   **Причина (в контексте VBA):** Требуется механизм для односторонней рассылки уведомлений от одного объекта многим, при этом список получателей может динамически меняться. Прямые вызовы методов от Субъекта к Наблюдателям создают сильную связанность.
*   **Ошибка (специфический симптом/пример):** Объект `CDataSource` после обновления данных должен уведомить несколько объектов `CChartUpdater`, `CTableRefresher`, `CNotificationPopup`. `CDataSource` содержит прямые ссылки на эти объекты и вызывает их методы `Update()`. При добавлении нового типа "слушателя" приходится изменять `CDataSource`.
*   **Решение (VBA-специфический паттерн/подход):**
    1.  Определить интерфейс Наблюдателя (`IObserver`) с методом `Update()`.
    2.  Конкретные Наблюдатели реализуют `IObserver`.
    3.  Субъект (`CSubject`) хранит коллекцию зарегистрированных Наблюдателей (объектов, реализующих `IObserver`).
    4.  Субъект предоставляет методы `Attach(obs As IObserver)` и `Detach(obs As IObserver)` для управления списком подписчиков.
    5.  Когда в Субъекте происходит значимое событие, он вызывает метод `Notify()`, который, в свою очередь, проходит по всем зарегистрированным Наблюдателям и вызывает у каждого метод `Update()`.
    *   **В VBA также можно использовать встроенные события (`Event` и `WithEvents`) для более простой реализации этого паттерна, особенно если Субъект и Наблюдатели находятся в одном проекте.** Однако, классический паттерн Observer более гибок для сценариев, где Наблюдатели могут быть из разных источников или требуют более сложного управления.
*   **Примеры кода:**

    **Классический Observer (более сложный, но гибкий):**
    ````vba
    ' --- Интерфейс Наблюдателя ---
    ' IObserver - Модуль класса
    Option Explicit
    Public Sub UpdateObserver(subjectState As Variant) ' Имя метода может быть любым, например, Refresh, OnNotify
    End Sub

    ' --- Конкретные Наблюдатели ---
    ' CConcreteObserverA - Модуль класса
    Option Explicit
    Implements IObserver
    Private m_Name As String
    Public Sub Initialize(name As String): m_Name = name: End Sub
    Private Sub IObserver_UpdateObserver(subjectState As Variant)
        Debug.Print "CConcreteObserverA (" & m_Name & "): Получено обновление. Новое состояние субъекта: " & CStr(subjectState)
    End Sub
    Private Sub Class_Initialize(): Debug.Print "CConcreteObserverA создан.": End Sub
    Private Sub Class_Terminate(): Debug.Print "CConcreteObserverA (" & m_Name & ") уничтожен.": End Sub

    ' CConcreteObserverB - Модуль класса
    Option Explicit
    Implements IObserver
    Private m_ID As Long
    Public Sub Setup(idNum As Long): m_ID = idNum: End Sub
    Private Sub IObserver_UpdateObserver(subjectState As Variant)
        Debug.Print "CConcreteObserverB (ID: " & m_ID & "): Уведомлен. Состояние: " & CStr(subjectState)
    End Sub
    Private Sub Class_Initialize(): Debug.Print "CConcreteObserverB создан.": End Sub
    Private Sub Class_Terminate(): Debug.Print "CConcreteObserverB (ID: " & m_ID & ") уничтожен.": End Sub

    ' --- Субъект (Издатель) ---
    ' CSubject - Модуль класса
    Option Explicit
    Private m_Observers As Collection
    Private m_State As String

    Private Sub Class_Initialize()
        Set m_Observers = New Collection
        m_State = "Начальное состояние"
        Debug.Print "CSubject: создан."
    End Sub

    Public Sub Attach(observer As IObserver)
        If observer Is Nothing Then Exit Sub
        ' Проверка на дубликаты (опционально, но полезно)
        Dim obs As IObserver
        For Each obs In m_Observers
            If obs Is observer Then Exit Sub ' Уже подписан
        Next
        m_Observers.Add observer
        Debug.Print "CSubject: Наблюдатель добавлен. Всего наблюдателей: " & m_Observers.Count
    End Sub

    Public Sub Detach(observer As IObserver)
        If observer Is Nothing Then Exit Sub
        Dim i As Long
        For i = 1 To m_Observers.Count
            If m_Observers(i) Is observer Then
                m_Observers.Remove i
                Debug.Print "CSubject: Наблюдатель удален. Осталось наблюдателей: " & m_Observers.Count
                Exit For
            End If
        Next
    End Sub

    Public Sub Notify()
        Debug.Print "CSubject: Уведомление наблюдателей..."
        Dim observer As IObserver
        If m_Observers.Count > 0 Then
            ' Создаем копию коллекции для итерации, если наблюдатели могут отписываться в своем методе Update
            Dim observersCopy As Collection
            Set observersCopy = New Collection
            For Each observer In m_Observers
                observersCopy.Add observer
            Next
            
            For Each observer In observersCopy
                observer.UpdateObserver m_State ' Вызов метода интерфейса
            Next
            Set observersCopy = Nothing
        Else
            Debug.Print "CSubject: Нет наблюдателей для уведомления."
        End If
    End Sub

    Public Property Let State(value As String)
        If m_State <> value Then
            m_State = value
            Debug.Print "CSubject: Состояние изменено на '" & m_State & "'."
            Notify ' Уведомить всех при изменении состояния
        End If
    End Property
    Public Property Get State() As String
        State = m_State
    End Property
    
    Private Sub Class_Terminate()
        Debug.Print "CSubject: уничтожается. Очистка наблюдателей..."
        Dim obs As IObserver
        ' Важно: если наблюдатели должны быть "живы" дольше субъекта,
        ' то простое Set m_Observers = Nothing может быть недостаточно,
        ' если субъект - единственное, что удерживает их.
        ' В данном случае, если наблюдатели не имеют других ссылок, они будут уничтожены.
        If Not m_Observers Is Nothing Then
            ' Не нужно вызывать Detach для каждого, просто обнуляем коллекцию.
            ' Ссылки из коллекции на наблюдателей исчезнут.
            Set m_Observers = Nothing
        End If
        Debug.Print "CSubject: уничтожен."
    End Sub

    ' --- Клиентский код ---
    ' Sub TestObserverPattern()
    '     Dim subject As CSubject
    '     Set subject = New CSubject
    '
    '     Dim obsA1 As CConcreteObserverA
    '     Set obsA1 = New CConcreteObserverA
    '     obsA1.Initialize "Первый Альфа"
    '
    '     Dim obsB1 As CConcreteObserverB
    '     Set obsB1 = New CConcreteObserverB
    '     obsB1.Setup 101
    '
    '     Dim obsA2 As CConcreteObserverA
    '     Set obsA2 = New CConcreteObserverA
    '     obsA2.Initialize "Второй Альфа"
    '
    '     subject.Attach obsA1
    '     subject.Attach obsB1
    '
    '     subject.State = "Новое состояние 1" ' obsA1 и obsB1 получат уведомление
    '
    '     subject.Attach obsA2
    '     subject.Detach obsA1 ' Отписываем первого
    '
    '     subject.State = "Финальное состояние" ' obsB1 и obsA2 получат уведомление
    '
    '     Set subject = Nothing ' Уничтожит субъект, что должно освободить ссылки на оставшихся наблюдателей
    '                         ' (если субъект был единственным, кто их держал)
    '     Set obsA1 = Nothing
    '     Set obsB1 = Nothing
    '     Set obsA2 = Nothing
    '     Debug.Print "--- Тест Observer завершен ---"
    ' End Sub
    ````
*   **Резюме:** Паттерн Наблюдатель (Observer) позволяет объектам подписываться на изменения в другом объекте и получать уведомления, не создавая жестких связей. Это способствует созданию гибких систем, где компоненты могут взаимодействовать, не зная о конкретных реализациях друг друга. В VBA также можно использовать встроенные события для схожих целей.
*   **Контрольные правила:**
    *   `Must`: Использовать паттерн Наблюдатель (или встроенные события VBA), когда один объект должен уведомлять множество других объектов об изменениях своего состояния, и список этих объектов может меняться динамически.
    *   `Must`: Определить четкий интерфейс для Наблюдателей.
    *   `Must`: Субъект должен предоставлять методы для подписки (`Attach`) и отписки (`Detach`) Наблюдателей.
    *   `Must Not`: Субъект не должен знать о конкретных классах Наблюдателей, а только об их общем интерфейсе.

### Резюме Раздела 6
Адаптация классических паттернов проектирования, таких как Фабрика и Наблюдатель, к специфике VBA может значительно улучшить структуру, гибкость и поддерживаемость приложений. Фабрика помогает инкапсулировать логику создания объектов, а Наблюдатель обеспечивает механизм слабой связи для уведомлений об изменениях состояния. Понимание и применение этих паттернов обогащает инструментарий VBA-разработчика.

### Золотые правила Раздела 6
1.  **Инкапсулируй создание:** Используй Фабрику для сложной или условной логики инстанцирования объектов.
2.  **Уведомляй, не зная:** Используй Наблюдатель (или события VBA) для оповещения зависимых объектов об изменениях, не создавая жестких связей.
3.  **Адаптируй, а не слепо копируй:** Понимай суть паттерна и адаптируй его реализацию к возможностям и ограничениям VBA.
4.  **Простота – ключ:** Не усложняй без необходимости. Если более простое решение (например, встроенные события VBA вместо полной реализации Observer) подходит, используй его.

### Мастер-класс Раздела 6: Система уведомлений о ценах акций (Фабрика + Наблюдатель)

**Сценарий:**
1.  `CStockTicker` (Субъект): Отслеживает цену акции. При изменении цены уведомляет подписчиков.
2.  `IStockObserver` (Интерфейс Наблюдателя): Определяет метод `PriceChanged`.
3.  `CSmsNotifier`, `CEmailNotifier` (Конкретные Наблюдатели): Реализуют `IStockObserver` и по-разному реагируют на изменение цены.
4.  `CNotifierFactory` (Фабрика): Создает экземпляры `CSmsNotifier` или `CEmailNotifier` по запросу.

**1. Интерфейс `IStockObserver`**
````vba
' IStockObserver - Модуль класса
Option Explicit
Public Sub PriceChanged(stockSymbol As String, newPrice As Double)
End Sub
````

**2. Конкретные Наблюдатели**
````vba
' CSmsNotifier - Модуль класса
Option Explicit
Implements IStockObserver
Private m_PhoneNumber As String

Public Sub Initialize(phoneNumber As String)
    m_PhoneNumber = phoneNumber
    Debug.Print "CSmsNotifier: Инициализирован для номера " & m_PhoneNumber
End Sub

Private Sub IStockObserver_PriceChanged(stockSymbol As String, newPrice As Double)
    Debug.Print "CSmsNotifier (" & m_PhoneNumber & "): АКЦИЯ " & stockSymbol & " ИЗМЕНИЛА ЦЕНУ! Новая цена: " & FormatCurrency(newPrice)
    ' Логика отправки SMS
End Sub
Private Sub Class_Terminate(): Debug.Print "CSmsNotifier (" & m_PhoneNumber & ") уничтожен.": End Sub

' CEmailNotifier - Модуль класса
Option Explicit
Implements IStockObserver
Private m_EmailAddress As String

Public Sub Initialize(emailAddress As String)
    m_EmailAddress = emailAddress
    Debug.Print "CEmailNotifier: Инициализирован для email " & m_EmailAddress
End Sub

Private Sub IStockObserver_PriceChanged(stockSymbol As String, newPrice As Double)
    Debug.Print "CEmailNotifier (" & m_EmailAddress & "): Цена акции " & stockSymbol & " теперь " & FormatCurrency(newPrice) & ". Отправка email..."
    ' Логика отправки Email
End Sub
Private Sub Class_Terminate(): Debug.Print "CEmailNotifier (" & m_EmailAddress & ") уничтожен.": End Sub
````

**3. Фабрика `CNotifierFactory`**
````vba
' CNotifierFactory - Модуль класса
Option Explicit

Public Function CreateNotifier(notifierType As String, destination As String) As IStockObserver
    Dim notifier As IStockObserver
    
    Select Case LCase(notifierType)
        Case "sms"
            Dim sms As CSmsNotifier
            Set sms = New CSmsNotifier
            sms.Initialize destination ' destination здесь - номер телефона
            Set notifier = sms
        Case "email"
            Dim email As CEmailNotifier
            Set email = New CEmailNotifier
            email.Initialize destination ' destination здесь - адрес email
            Set notifier = email
        Case Else
            Debug.Print "CNotifierFactory: Неизвестный тип уведомителя: " & notifierType
            Set notifier = Nothing
    End Select
    
    Set CreateNotifier = notifier
End Function
Private Sub Class_Initialize(): Debug.Print "CNotifierFactory: создана.": End Sub
Private Sub Class_Terminate(): Debug.Print "CNotifierFactory: уничтожена.": End Sub
````

**4. Субъект `CStockTicker`**
````vba
' CStockTicker - Модуль класса
Option Explicit
Private m_Symbol As String
Private m_CurrentPrice As Double
Private m_Observers As Collection

Private Sub Class_Initialize()
    Set m_Observers = New Collection
    Debug.Print "CStockTicker: создан."
End Sub

Public Sub InitializeTicker(stockSymbol As String, initialPrice As Double)
    m_Symbol = stockSymbol
    m_CurrentPrice = initialPrice
    Debug.Print "CStockTicker: Инициализирован для " & m_Symbol & " с ценой " & FormatCurrency(m_CurrentPrice)
End Sub

Public Sub Attach(observer As IStockObserver)
    If observer Is Nothing Then Exit Sub
    m_Observers.Add observer
    Debug.Print "CStockTicker (" & m_Symbol & "): Подписчик добавлен. Всего: " & m_Observers.Count
End Sub

Public Sub Detach(observer As IStockObserver)
    If observer Is Nothing Then Exit Sub
    Dim i As Long
    For i = 1 To m_Observers.Count
        If m_Observers(i) Is observer Then
            m_Observers.Remove i
            Debug.Print "CStockTicker (" & m_Symbol & "): Подписчик удален. Осталось: " & m_Observers.Count
            Exit For
        End If
    Next
End Sub

Private Sub Notify()
    Debug.Print "CStockTicker (" & m_Symbol & "): Уведомление подписчиков об изменении цены до " & FormatCurrency(m_CurrentPrice)
    Dim obs As IStockObserver
    For Each obs In m_Observers
        obs.PriceChanged m_Symbol, m_CurrentPrice
    Next
End Sub

Public Property Let Price(value As Double)
    If Abs(m_CurrentPrice - value) > 0.001 Then ' Сравнение с допуском для Double
        m_CurrentPrice = value
        Debug.Print "CStockTicker (" & m_Symbol & "): Цена изменена на " & FormatCurrency(m_CurrentPrice)
        Notify
    End If
End Property
Public Property Get Price() As Double
    Price = m_CurrentPrice
End Property
Public Property Get Symbol() As String
    Symbol = m_Symbol
End Property

Private Sub Class_Terminate()
    Debug.Print "CStockTicker (" & m_Symbol & "): уничтожается. Очистка подписчиков..."
    Set m_Observers = Nothing ' Освобождает ссылки на подписчиков
    Debug.Print "CStockTicker (" & m_Symbol & ") уничтожен."
End Sub
````

**5. Стандартный модуль `StockMarketDemo`**
````vba
' Module: StockMarketDemo
Option Explicit

Sub RunStockNotificationSystem()
    ' 1. Создаем фабрику уведомителей
    Dim notifierFactory As CNotifierFactory
    Set notifierFactory = New CNotifierFactory
    
    ' 2. Создаем уведомителей через фабрику
    Dim smsAlert As IStockObserver
    Set smsAlert = notifierFactory.CreateNotifier("sms", "+1234567890")
    
    Dim emailAlert As IStockObserver
    Set emailAlert = notifierFactory.CreateNotifier("email", "user@example.com")

    Dim anotherSmsAlert As IStockObserver
    Set anotherSmsAlert = notifierFactory.CreateNotifier("sms", "+0987654321")
    
    ' 3. Создаем тикеры акций
    Dim acmeTicker As CStockTicker
    Set acmeTicker = New CStockTicker
    acmeTicker.InitializeTicker "ACME", 100.50
    
    Dim xyzTicker As CStockTicker
    Set xyzTicker = New CStockTicker
    xyzTicker.InitializeTicker "XYZ", 75.20
    
    ' 4. Подписываем уведомителей на тикеры
    acmeTicker.Attach smsAlert
    acmeTicker.Attach emailAlert
    
    xyzTicker.Attach emailAlert ' Email будет получать уведомления от обеих акций
    xyzTicker.Attach anotherSmsAlert
    
    ' 5. Имитируем изменения цен
    Debug.Print vbCrLf & "--- ИЗМЕНЕНИЕ ЦЕН ACME ---"
    acmeTicker.Price = 101.75 ' smsAlert и emailAlert получат уведомление
    
    Debug.Print vbCrLf & "--- ИЗМЕНЕНИЕ ЦЕН XYZ ---"
    xyzTicker.Price = 74.90  ' emailAlert и anotherSmsAlert получат уведомление
    
    ' 6. Отписываем одного уведомителя
    Debug.Print vbCrLf & "--- ОТПИСКА SMS от ACME ---"
    acmeTicker.Detach smsAlert
    acmeTicker.Price = 102.00 ' Только emailAlert получит уведомление от ACME
    
    ' 7. Очистка
    Debug.Print vbCrLf & "--- ЗАВЕРШЕНИЕ И ОЧИСТКА ---"
    Set acmeTicker = Nothing
    Set xyzTicker = Nothing
    
    Set smsAlert = Nothing
    Set emailAlert = Nothing
    Set anotherSmsAlert = Nothing
    
    Set notifierFactory = Nothing
    
    Debug.Print "--- Демо StockMarket завершено ---"
End Sub
````

**Как это работает:**
1.  `CNotifierFactory` создает различные типы уведомителей (`CSmsNotifier`, `CEmailNotifier`), реализующих интерфейс `IStockObserver`. Клиентский код не знает о конкретных классах уведомителей.
2.  `CStockTicker` (Субъект) отслеживает цену акции. Другие объекты могут подписаться на уведомления об изменении цены, реализуя `IStockObserver` и используя метод `Attach` тикера.
3.  Когда цена акции в `CStockTicker` изменяется (через свойство `Price`), он вызывает свой метод `Notify`, который проходит по всем подписчикам и вызывает у них метод `PriceChanged`.
4.  Это позволяет гибко добавлять или удалять подписчиков и типы уведомителей, не изменяя код `CStockTicker` или код, который инициирует изменение цены.

Этот мастер-класс демонстрирует совместное использование Фабрики для создания объектов-наблюдателей и паттерна Наблюдатель для организации системы уведомлений.

---
## Общее заключение

Проектирование архитектуры VBA-приложений, несмотря на некоторые ограничения языка, может быть выполнено на высоком профессиональном уровне. Ключ к успеху лежит в понимании специфики VBA и адаптации общих принципов и паттернов проектирования.

В первой части мы рассмотрели основы **инициализации**, важность четких **точек входа** и стратегии, такие как **Bootstrapping**, **Singleton** и **Lazy Initialization**. Также была затронута организация доступа к функциональности через правильное использование **стандартных модулей и модулей классов**, а также паттерна **Фасад**.

Во второй части основное внимание было уделено **управлению жизненным циклом объектов**, включая корректное создание и освобождение, понимание области видимости и времени жизни переменных, а также критически важный аспект избегания **циклических ссылок**. Затем мы погрузились в техники **внедрения зависимостей (DI)** – "Constructor" Injection, Property Injection и Parameter Injection, а также рассмотрели Service Locator, отметив его потенциальные недостатки. Наконец, были представлены адаптированные для VBA паттерны **Фабрика** и **Наблюдатель**, которые помогают решать распространенные задачи проектирования.

Применение этих концепций и паттернов позволяет создавать VBA-приложения, которые являются:
*   **Более надежными:** Благодаря контролируемой инициализации и управлению жизненным циклом.
*   **Более гибкими и расширяемыми:** За счет слабой связанности, достигаемой через DI и паттерны.
*   **Более тестируемыми:** Возможность подмены зависимостей упрощает создание юнит-тестов.
*   **Более понятными и поддерживаемыми:** Четкая структура и разделение ответственности облегчают понимание и внесение изменений.

Для ИИ-ассистентов, генерирующих VBA-код, крайне важно не просто воспроизводить синтаксис, а следовать этим архитектурным принципам. Это позволит создавать решения, которые не только работают, но и являются качественными с точки зрения инженерии программного обеспечения.

---

## Список использованных источников (иллюстративный, Часть 2)

*Как и в первой части, это иллюстративный список типов источников.*

1.  **Официальная документация Microsoft (MSDN/Docs):**
    *   Статьи о подсчете ссылок в COM (RefCount).
    *   Примеры использования `Class_Initialize` и `Class_Terminate`.
    *   Документация по событиям VBA (`Event`, `WithEvents`, `RaiseEvent`).
    *   *Ключевые моменты:* Основы управления памятью, события, область видимости переменных.
2.  **Специализированные форумы и сообщества (StackOverflow, RubberduckVBA, etc.):**
    *   Обсуждения утечек памяти в VBA и циклических ссылок.
    *   Примеры реализации DI-контейнеров или подходов к DI в VBA.
    *   Адаптация паттернов проектирования (Observer, Factory, etc.) для VBA.
    *   *Ключевые моменты:* Практические проблемы и их решения, опыт сообщества.
3.  **Книги по разработке ПО и паттернам проектирования (с адаптацией для VBA):**
    *   Fowler, Martin. "Patterns of Enterprise Application Architecture." (Концепции Service Locator, DI).
    *   Fowler, Martin. "Refactoring: Improving the Design of Existing Code." (Принципы улучшения кода).
    *   Gamma, E., et al. "Design Patterns: Elements of Reusable Object-Oriented Software." (Классические описания паттернов).
    *   *Ключевые моменты:* Теоретические основы DI, паттерны, лучшие практики.
4.  **Блоги и статьи опытных VBA-разработчиков:**
    *   Статьи, посвященные ООП в VBA, управлению памятью, DI.
    *   Примеры реализации сложных систем на VBA с хорошей архитектурой.
    *   *Ключевые моменты:* Глубокий анализ специфики VBA, практические советы по архитектуре.
5.  **Ресурсы по VB6 (часто применимы к VBA с небольшими изменениями):**
    *   Старые статьи и книги по объектно-ориентированному программированию на VB6.
    *   Обсуждения управления памятью и COM в VB6.
    *   *Ключевые моменты:* Схожие концепции COM, управление объектами.


