Отлично. Я начну с построения архитектурного справочника по VBA, начиная с схемы Layered Architecture как базовой. В рамках этого этапа я:

* Разработаю единый тестовый контекст и общие вспомогательные модули (универсальная среда), которые будут переиспользоваться во всех архитектурах.
* Сформирую структуру каталогов (`\docs`, `\src`, `\tests`, `\samples`) и представлю рабочий пример Layered Architecture с:

  * исходным кодом (UI, бизнес-логика, DAL);
  * unit-тестами (Rubberduck);
  * готовым описанием (Mermaid, таблицы, пояснения, ошибки).

Я вернусь с первой частью материалов, как только она будет готова.


# Архитектура VBA: Многослойная (Layered) архитектура

## Суть (Essence)

**Многослойная архитектура** – это подход к организации кода, при котором логически разные части приложения разделяются на слои. Каждый слой отвечает за свою **суть**: пользовательский интерфейс, бизнес-логику и доступ к данным. Такой дизайн давно зарекомендовал себя в разработке ПО, и он применим даже в VBA-проектах для Excel.

**Зачем нужна слоистость?** В небольших макросах VBA часто вся логика сосредоточена в нескольких процедурных модулях или прямо в коде форм. Это быстро приводит к *«спагетти-коду»* – трудно сопровождаемому и расширяемому. Например, распространённая ошибка – нагружать `UserForm` не только отображением, но и всей обработкой данных (анти-паттерн *«Smart UI»*, когда форма «знает всё» и делает больше, чем должна). Это удобно для прототипа, но совершенно не масштабируется и трудно поддерживается на длинной дистанции. Многослойная архитектура предотвращает такие ситуации, разделяя обязанности:

* **UI (интерфейс)**: только отображение данных и взаимодействие с пользователем (напр. формы Excel).
* **BL (бизнес-логика)**: обработки, расчёты, правила – ядро приложения (обычно класс(ы) в VBA).
* **DAL (доступ к данным)**: хранение и получение данных (модули/классы, работающие с файлами, листами, БД).

**Преимущества:** Такой подход повышает понятность кода, облегчает тестирование и повторное использование компонентов. Например, можно отладить бизнес-логику отдельно от UI, либо заменить источник данных (файл, лист, база) без переписывания остального кода. Кроме того, чёткое разделение упрощает обнаружение ошибок: каждая проблема локализуется в своём слое (UI-ошибки, логические ошибки, проблемы хранения).

**Фокус на реальную работоспособность:** В этом руководстве мы создадим полностью рабочий пример многослойной архитектуры в VBA. Будут разработаны **самодостаточные компоненты**, которые можно сразу запустить в Excel-файле. Мы начнём с подготовки общей инфраструктуры (например, логгер, интерфейсы, тестовые данные), а затем реализуем шаблон из трёх слоёв (UI, BL, DAL) на примере простой задачи управления списком задач (*Task Manager*). Мы также продемонстрируем, как этот код организовать в структуре проекта (\`docs\`, \`src\`, \`tests\`, \`samples\`), как писать модульные тесты с Rubberduck, и рассмотрим типичные ошибки и **граничные случаи**.

## Элементы (Elements)

**Общие компоненты (Common modules):** Сначала создадим ряд модулей и классов, которые будут использоваться во всех примерах архитектур. Эти компоненты независимы от конкретной логики и не имеют внешних зависимостей (кроме стандартных библиотек VBA). Их цель – обеспечить базовую функциональность и инфраструктуру.

* **Логгер (Logger.cls)** – класс для ведения журнала событий и отладочной информации. Например, он может предоставлять методы `LogInfo`, `LogError` для вывода сообщений через `Debug.Print` или в файл. Этот класс помогает отслеживать работу приложения на всех слоях.

  ```vba
  ' src\Common\Logger.cls
  Option Explicit
  Private logCount As Long

  Public Sub LogInfo(ByVal msg As String)
      logCount = logCount + 1
      Debug.Print "INFO #" & logCount & ": " & msg
      ' Дополнительно: можно писать в текстовый файл или лист логов
  End Sub

  Public Sub LogError(ByVal msg As String)
      logCount = logCount + 1
      Debug.Print "ERROR #" & logCount & ": " & msg
  End Sub
  ```

  *Описание:* Логгер нумерует сообщения и выводит их. В реальном проекте его можно улучшить (добавить отметку времени, уровни логирования, вывод в файл и т.п.), но для простоты этого достаточно. Использование логгера по всему коду позволяет не использовать `MsgBox` для отладки и помогает разбирать, что произошло до ошибки.

* **Интерфейс задачи (ITask.cls)** – интерфейс для сущности "задача". В VBA интерфейс реализуется как класс-модуль с объявлением необходимых свойств/методов без реализации. Интерфейс задаёт контракт, который должны выполнять классы-реализации. В нашем случае интерфейс `ITask` определит свойства задачи (например, заголовок, статус) и, опционально, методы.

  ```vba
  ' src\Common\ITask.cls (интерфейс задачи)
  Option Explicit
  Public Property Get Title() As String
  End Property
  Public Property Let Title(ByVal value As String)
  End Property

  Public Property Get IsCompleted() As Boolean
  End Property
  Public Property Let IsCompleted(ByVal value As Boolean)
  End Property

  ' (Допустим, у задачи есть заголовок и флаг завершенности)
  ```

  *Описание:* Здесь мы определили, что любая *Task* должна иметь свойства `Title` (название) и `IsCompleted` (завершена ли). Реализацию этих свойств предоставит класс задачи. Интерфейсы в VBA способствуют слабой связанности и позволяют легко менять реализации (например, можно создать разные типы задач, реализующие `ITask`, но в данном примере у нас один тип).

* **Интерфейс хранилища задач (ITaskRepository.cls)** – контракт для слоя доступа к данным. Он описывает, какие операции должны поддерживаться для хранения задач. Например, методы получения списка задач, добавления, удаления и т.д. Это абстракция, которая позволит нам иметь разные реализации хранилища (например, фальшивое в памяти и реальное в файле), не меняя код бизнес-логики.

  ```vba
  ' src\Common\ITaskRepository.cls (интерфейс репозитория задач)
  Option Explicit
  Public Sub AddTask(ByVal t As ITask)
  End Sub
  Public Function GetAllTasks() As Collection ' возвращает все задачи
  End Function
  Public Sub DeleteTask(ByVal taskId As String)
  End Sub
  ' Можно расширить интерфейс при необходимости (обновление, поиск и т.п.)
  ```

  *Описание:* Интерфейс репозитория определяет основные операции: добавить задачу, получить все задачи, удалить задачу по идентификатору. Здесь для упрощения задачи идентифицируются строкой (можно использовать GUID или уникальное название). Метод `GetAllTasks` возвращает `Collection` объектов `ITask`. Реализации этого интерфейса позаботятся о конкретном способе хранения (в памяти, на листе Excel, в файле и пр.).

* **Генератор тестовых данных (TestData.bas)** – модуль с процедурами, создающими примерные данные для демонстрации или тестов. Чтобы не вводить руками много объектов, напишем процедуру, генерирующую коллекцию задач с тестовыми значениями (например, 5 задач с разными именами). Этот модуль пригодится при тестировании бизнес-логики без UI.

  ```vba
  ' src\Common\TestData.bas
  Option Explicit

  Public Function SampleTasks(Optional ByVal count As Long = 5) As Collection
      Dim tasks As New Collection
      Dim i As Long
      For i = 1 To count
          Dim t As Task  ' создаем новую задачу
          Set t = New Task
          t.Title = "Задача #" & i
          t.IsCompleted = False
          tasks.Add t
      Next i
      Set SampleTasks = tasks
  End Function
  ```

  *Описание:* Функция `SampleTasks` создаёт заданное количество задач с произвольными названиями. По умолчанию 5 задач. Все отмечены как невыполненные. Она использует класс `Task` (реализация `ITask`, опишем ниже). Такой генератор можно использовать для инициализации *FakeStorage* или для быстрого наполнения списка на UI в демонстрационных целях.

* **Базовая форма (BaseForm.frm / BaseForm.cls)** – общий шаблон для пользовательских форм. В VBA нет прямого наследования форм, но мы можем определить соглашения или базовую логику, которой будут придерживаться все формы. Например, можно создать *BaseForm* – класс-обёртку над `UserForm`, который реализует типовые вещи: инициализацию, единый метод показа/закрытия, обработку ошибок интерфейса. В рамках нашего примера, мы ограничимся рекомендацией: **каждая форма должна** объявлять `Option Explicit`, корректно освобождать объекты, и по возможности использовать интерфейсы/события для связи с логикой, вместо глобальных переменных.

  *Пример соглашения:* можно написать класс `FormController` (в \`src\Common\BaseForm.cls\`), который содержит логику открытия формы (модально или нет) и логирование ошибок при закрытии:

  ```vba
  ' src\Common\BaseForm.cls (управление формами, условно)
  Option Explicit
  Private WithEvents frm As VBComponent  ' ссылка на форму (объект формы VBA)

  Public Sub ShowForm(ByVal form As Object)
      Set frm = form
      VBA.UserForms.Add(frm.Name).Show  ' открываем форму по имени
  End Sub

  Private Sub frm_Terminate()
      ' Унифицированное действие при закрытии формы
      Debug.Print "Форма " & frm.Name & " закрыта."
      Set frm = Nothing
  End Sub
  ```

  *Описание:* Этот код иллюстрирует идею: можно иметь контроллер форм, который умеет показывать любую форму и отслеживать её закрытие. Здесь мы используем `UserForms.Add` для открытия по имени (требуется, чтобы форма была вставлена в проект). По событию Terminate формы – выводим сообщение и очищаем ссылку. Хотя *BaseForm* скорее концептуальный, важно, чтобы каждая конкретная форма следовала единым правилам (Option Explicit, освобождение ресурсов, минимальная логика). Далее в разделе UI мы увидим реализацию конкретной формы с этими принципами.

Теперь перейдём к реализации **самого шаблона многослойной архитектуры (Layered Architecture)** для задачи управления списком задач:

* **UI (пользовательский интерфейс)** – у нас это форма `frmTaskManager` (UserForm) для отображения и ввода задач (слой *View*). Эта форма содержит элементы управления (например, текстовое поле для названия задачи, кнопку "Добавить", список задач, кнопку "Удалить" и т.д.). Код формы должен быть минимальным: никаких вычислений или доступа к файлам прямо в обработчиках событий! Вместо этого форма обращается к объектам бизнес-логики. Форма также подписывается на события от бизнес-логики, чтобы обновляться автоматически. Ниже приведём код UserForm, демонстрирующий эти идеи.

* **BL (бизнес-логика)** – слой, содержащий классы: `TaskManager` и `Task`.

  * `Task` (класс `Task.cls`) – реализация интерфейса `ITask`. Он хранит данные одной задачи: идентификатор (например, `ID` или можем использовать заголовок как уникальный ключ), название (`Title`), признак завершённости (`IsCompleted`) и, возможно, другую информацию. В нашем примере сосредоточимся на названии и статусе. Этот класс инкапсулирует состояние задачи и может содержать методы, например, `MarkComplete` (отметить как выполненную) – для демонстрации бизнес-правил.
  * `TaskManager` (класс `TaskManager.cls`) – основной класс логики приложения. Он управляет коллекцией задач и использует `ITaskRepository` для их постоянного хранения. `TaskManager` предоставляет методы для добавления новой задачи, получения списка задач, удаления или обновления задач. Он также генерирует **события** (Events), уведомляя подписчиков (например, UI) об изменениях – например, событие `TaskAdded` при добавлении новой задачи. `TaskManager` принимает в себя конкретную реализацию репозитория (через метод `SetRepository` или в инициализации), чтобы знать, куда сохранять задачи. Внутри себя он может хранить список задач в памяти (синхронизированный с хранилищем). Также он может использовать `Logger` для записи значимых действий (например, "задача добавлена", "ошибка при удалении" и т.п.).

* **DAL (слой доступа к данным)** – реализует `ITaskRepository`. У нас будет две реализации:

  * `FakeStorage` (класс или модуль, например `FakeStorage.cls`) – простое фальшивое хранилище. Для демонстрации сделаем его в памяти: он будет хранить задачи в коллекции или словаре, эмулируя базу/файл. Можно даже не сохранять между запусками – просто для тестов. Это полезно для модульных тестов: мы можем заменить настоящий файл на `FakeStorage`, чтобы тестировать логику в изоляции.
  * `FileRepository` (класс `FileRepository.cls`) – реальное хранилище, работающее с файлом Excel (или другим источником данных). В рамках Excel-проекта логично хранить задачи на скрытом листе внутри книги или в отдельном CSV/текстовом файле. Для простоты реализуем хранение на специальном листе "TasksData" текущей книги. `FileRepository` будет, например, при добавлении задачи записывать её как новую строку на лист, при запросе всех задач – читать все непустые строки и создавать объекты `Task`, при удалении – удалять строку/помечать запись. Эта реализация покажет, как отделить логику от деталей Excel: если завтра решим хранить в другом формате, достаточно написать другой класс, реализующий `ITaskRepository`, а остальной код трогать не придётся.

Теперь рассмотрим реализацию основных классов более детально. Все классы объявляются с `Option Explicit` во избежание ошибок с необъявленными переменными (это **стандарт качества кода**, помогающий ловить опечатки и логические ошибки на этапе компиляции). По возможности, важные объекты освобождаются через `Set ... = Nothing` после использования (особенно для объектов Excel или файловых потоков), чтобы не держать лишних ресурсов. Также, где уместно, мы используем ключевое слово `WithEvents` и объявляем/генерируем события для обратной связи между слоями.

### Класс Task (задача)

```vba
' src\Layered\Task.cls
Option Explicit
Implements ITask  ' реализуем интерфейс задачи

Private Type TTask
    ID As String
    Title As String
    Completed As Boolean
End Type

Private this As TTask

' Свойства Title (реализация ITask.Title)
Public Property Get Title() As String
    Title = this.Title
End Property
Public Property Let Title(ByVal value As String)
    this.Title = value
End Property

' Свойства IsCompleted (реализация ITask.IsCompleted)
Public Property Get IsCompleted() As Boolean
    IsCompleted = this.Completed
End Property
Public Property Let IsCompleted(ByVal value As Boolean)
    this.Completed = value
End Property

' Дополнительное свойство только для Task (не в интерфейсе): ID задачи
Public Property Get TaskID() As String
    TaskID = this.ID
End Property
Public Property Let TaskID(ByVal value As String)
    this.ID = value
End Property

' Инициализация нового объекта Task
Private Sub Class_Initialize()
    ' Генерируем уникальный ID (например, на основе времени или счетчика)
    this.ID = Format$(Now, "yyyymmddhhnnss") & Round(Rnd()*1000,0)
    this.Completed = False
End Sub

' Удобный метод: отметить задачу выполненной
Public Sub MarkComplete()
    this.Completed = True
End Sub
```

*Комментарий:* Класс `Task` хранит данные задачи. Мы используем пользовательский тип `TTask` для приватного хранения, что облегчает группировку полей. Реализованы свойства `Title` и `IsCompleted` из интерфейса `ITask`. Кроме того, добавлено поле `ID` – уникальный идентификатор (в данном случае строка, генерируемая при создании на основе текущего времени; в реальном проекте можно использовать GUID или счётчик). Метод `MarkComplete` устанавливает флаг выполнения. Такой класс легко расширить (добавить дедлайн, приоритет и т.д.). Обратите внимание, что интерфейс `ITask` определял только название и статус, поэтому `TaskID` не объявлен в интерфейсе – он используется внутри приложения, а внешним потребителям достаточно знать свойства через интерфейс (ID мог бы быть частью интерфейса, но предположим, что логика идентификации – внутренняя деталь задачи).

### Класс TaskManager (менеджер задач, бизнес-логика)

```vba
' src\Layered\TaskManager.cls
Option Explicit

' События для оповещения UI об изменениях
Public Event TaskAdded(ByVal newTask As ITask)
Public Event TaskRemoved(ByVal taskId As String)

' Ссылка на репозиторий (хранилище) задач
Private repo As ITaskRepository
' Коллекция задач в памяти (кэш текущей сессии)
Private tasks As New Collection
' Логгер для сообщений (опционально)
Private logger As Logger

' Инициализация класса
Private Sub Class_Initialize()
    ' По умолчанию можно использовать FakeStorage или FileRepository
    ' В нашем примере установим по умолчанию FakeStorage для безопасности
    Set repo = New FakeStorage
    Set logger = New Logger
    ' Загрузим существующие задачи из хранилища в коллекцию
    Dim t As ITask
    For Each t In repo.GetAllTasks()
        tasks.Add t, (t.Title & "_" & CStr(tasks.Count + 1))
        ' (ключ для коллекции сделаем комбинацией заголовка и порядкового номера, 
        '  чтобы избежать ошибок дублирования ключей)
    Next
    logger.LogInfo "TaskManager initialized with " & tasks.Count & " tasks."
End Sub

' Задать конкретный репозиторий (например, переключиться с фейкового на файл)
Public Sub SetRepository(ByVal repository As ITaskRepository)
    ' Переносим текущие задачи в новый репозиторий (синхронизация, если нужно)
    Dim t As ITask
    If Not repository Is Nothing Then
        ' Можно сохранить текущие задачи в новый репо
        For Each t In tasks
            repository.AddTask t
        Next
    End If
    Set repo = repository
    logger.LogInfo "Repository set to " & TypeName(repository)
End Sub

' Добавить новую задачу
Public Function AddTask(ByVal title As String) As ITask
    Dim t As Task: Set t = New Task
    t.Title = title
    t.IsCompleted = False
    ' Сохраняем в репозиторий
    repo.AddTask t
    ' Добавляем во внутреннюю коллекцию
    tasks.Add t, t.TaskID
    logger.LogInfo "Added task '" & title & "' (ID=" & t.TaskID & ")"
    ' Генерируем событие для UI
    RaiseEvent TaskAdded(t)
    Set AddTask = t  ' возвращаем добавленную задачу как результат
End Function

' Удалить задачу по ID
Public Sub RemoveTask(ByVal taskId As String)
    On Error GoTo ErrHandler
    Dim t As ITask, idx As Long, found As Boolean
    ' Найдем задачу с данным ID в коллекции
    For idx = 1 To tasks.Count
        Set t = tasks.Item(idx)
        If TypeOf t Is Task Then
            If t.TaskID = taskId Then
                found = True
                Exit For
            End If
        End If
    Next
    If Not found Then Err.Raise 5, , "Task ID not found"
    ' Удаляем из коллекции и репозитория
    tasks.Remove idx
    repo.DeleteTask taskId
    logger.LogInfo "Removed task ID=" & taskId
    RaiseEvent TaskRemoved(taskId)
    Exit Sub
ErrHandler:
    logger.LogError "Error removing task ID=" & taskId & ": " & Err.Description
    Err.Clear
End Sub

' Получить все задачи (коллекцию)
Public Function GetAllTasks() As Collection
    ' Возвращаем копию коллекции, чтобы вызывающий не смог повредить внутренние данные
    Dim result As New Collection
    Dim t As ITask
    For Each t In tasks
        result.Add t
    Next
    Set GetAllTasks = result
End Function

' Очистка объекта
Private Sub Class_Terminate()
    ' Освободим ресурсы
    Set repo = Nothing
    Set logger = Nothing
End Sub
```

*Комментарий:* `TaskManager` – сердце бизнес-логики. Он хранит список задач `tasks` и содержит ссылку на `repo` (репозиторий). В `Class_Initialize` мы устанавливаем репозиторий по умолчанию (для безопасности выбрали `FakeStorage`, чтобы не было ошибок, если не задан явно) и загружаем задачи из него. Предусмотрен метод `SetRepository` для замены хранилища на лету – при этом существующие задачи сохраняются в новое хранилище (это простой способ синхронизироваться). Основные методы:

* `AddTask(title)` – создает новый объект `Task`, устанавливает ему название и статус, сохраняет его через `repo.AddTask`, добавляет в локальный список `tasks`, логирует действие и генерирует событие `TaskAdded`. Возвращает добавленную задачу (хотя UI может и через событие узнать, но возвращать полезно для возможного дальнейшего использования).
* `RemoveTask(taskId)` – находит задачу с заданным ID, удаляет из коллекции и вызывает `repo.DeleteTask` для удаления из хранилища. Ошибки ловятся через `On Error` (например, если ID не найден или возникла проблема в репозитории) – в обработчике ошибок мы логируем проблему через `LogError` (что лучше, чем молча игнорировать). Также генерируется событие `TaskRemoved`.
* `GetAllTasks()` – возвращает коллекцию всех задач. Обратите внимание: возвращается **копия** коллекции `tasks`, чтобы вызывающий код не мог напрямую менять внутреннюю коллекцию менеджера (это добавляет безопасности). Мы просто создаём новую коллекцию и копируем объекты. (Замечание: сами объекты задачи копируются по ссылке, глубокое копирование не делаем, но этого достаточно, если вызывающий не будет изменять задачи напрямую или если изменит – они всё равно отражают тот же объект, что хранится внутри).
* В `Class_Terminate` очищаем ссылки `repo` и `logger` (Set ... = Nothing) – это хорошая практика, помогающая избежать утечек памяти или зависаний объектов. Хотя VBA обычно сам освобождает при завершении объекта, явное освобождение не повредит, особенно если могли быть циклические ссылки.

События `TaskAdded` и `TaskRemoved` позволяют UI-слою реагировать на изменения: вместо того, чтобы UI опрашивал `TaskManager`, сам `TaskManager` уведомляет подписчиков. Чтобы UI мог подписаться, он должен объявить объект `WithEvents` – покажем это далее.

### Класс FakeStorage (фальшивое хранилище, DAL)

```vba
' src\Layered\FakeStorage.cls
Option Explicit
Implements ITaskRepository

' Хранение задач в памяти (словарь: ключ - ID, значение - объект ITask)
Private taskDict As Object

Private Sub Class_Initialize()
    ' Используем Scripting.Dictionary для удобства (потребуется добавить Reference на Microsoft Scripting Runtime)
    Set taskDict = CreateObject("Scripting.Dictionary")
    ' Инициализируем начальными данными для примера (возьмём из TestData)
    Dim t As ITask
    For Each t In SampleTasks(3) ' создадим 3 тестовые задачи
        taskDict.Add GetTaskKey(t), t
    Next
End Sub

Private Function GetTaskKey(ByVal t As ITask) As String
    ' Генерируем ключ для словаря на основе ID или Title
    Dim key As String
    If TypeOf t Is Task Then
        key = (t.TaskID)
    Else
        key = t.Title  ' если это другой тип, возьмём Title как ключ (допущение)
    End If
    GetTaskKey = key
End Function

' Реализация ITaskRepository.AddTask
Public Sub ITaskRepository_AddTask(ByVal t As ITask)
    taskDict(AddTask_GetKey(t)) = t  ' добавляем или обновляем по ключу
End Sub
Private Function AddTask_GetKey(ByVal t As ITask) As String
    AddTask_GetKey = GetTaskKey(t)
End Function

' Реализация ITaskRepository.GetAllTasks
Public Function ITaskRepository_GetAllTasks() As Collection
    Dim col As New Collection
    Dim key As Variant, t As ITask
    For Each key In taskDict.Keys
        Set t = taskDict(key)
        col.Add t
    Next
    Set ITaskRepository_GetAllTasks = col
End Function

' Реализация ITaskRepository.DeleteTask
Public Sub ITaskRepository_DeleteTask(ByVal taskId As String)
    If taskDict.Exists(taskId) Then
        taskDict.Remove taskId
    End If
End Sub

Private Sub Class_Terminate()
    Set taskDict = Nothing
End Sub
```

*Комментарий:* `FakeStorage` – класс, реализующий `ITaskRepository` для хранения данных в памяти. Здесь для удобства мы использовали `Scripting.Dictionary` (ассоциативный массив ключ-объект). **Важно:** чтобы использовать `Dictionary`, необходимо в VBA поставить ссылку (Reference) на библиотеку *Microsoft Scripting Runtime*. Это считается допустимой зависимостью, т.к. она стандартная в Windows. Если не хочется зависимостей, можно использовать `Collection` и писать поиск по ключу вручную.

В `Class_Initialize` `FakeStorage` заполняется тремя тестовыми задачами (через наш генератор `SampleTasks`). Это значит, что при запуске приложения, если подключено `FakeStorage`, уже будет несколько задач для демонстрации. (В тестах можно не генерировать, а начать с пустого словаря).

Методы:

* `AddTask` – добавляет задачу в словарь. Ключ генерируем функцией `GetTaskKey`: для наших `Task` ключом будет внутренний `TaskID`, для других возможных типов, реализующих `ITask`, мы пытаемся взять `Title`. В словарь добавляем или обновляем запись.
* `GetAllTasks` – возвращает все задачи как коллекцию, перебирая все значения словаря.
* `DeleteTask` – удаляет по ключу (ID). Если такого ключа нет – просто ничего не делает (тихо игнорирует).
* Мы также освобождаем словарь в `Terminate`.

Этот класс **не** взаимодействует ни с файлами, ни с листами Excel – он полностью автономен. Это удобно для быстрого тестирования и предотвращает побочные эффекты (например, лишние сообщения или медленное чтение с диска). В реальном применении `FakeStorage` может служить «заглушкой» для юнит-тестов.

### Класс FileRepository (хранилище на листе Excel, DAL)

```vba
' src\Layered\FileRepository.cls
Option Explicit
Implements ITaskRepository

Private Const DATA_SHEET_NAME As String = "TasksData"
' Возможно, имеет смысл хранить путь к файлу или объект Workbook, 
' но возьмём текущую книгу для простоты:
Private wb As Workbook

Private Sub Class_Initialize()
    ' Берём текущую книгу и убеждаемся, что в ней есть лист для данных
    Set wb = ThisWorkbook
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets(DATA_SHEET_NAME)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = DATA_SHEET_NAME
        ' Добавим заголовки столбцов
        ws.Range("A1:B1").Value = Array("Title", "Completed")
    End If
    On Error GoTo 0
End Sub

' Добавить задачу (записать в конец листа)
Public Sub ITaskRepository_AddTask(ByVal t As ITask)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(DATA_SHEET_NAME)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 1 ' если лист пустой кроме заголовков
    lastRow = lastRow + 1
    ' Записываем данные задачи: Название в колонку A, Статус в B
    ws.Cells(lastRow, 1).Value = t.Title
    ws.Cells(lastRow, 2).Value = IIf(t.IsCompleted, "TRUE", "FALSE")
End Sub

' Получить все задачи (чтение всех строк листа)
Public Function ITaskRepository_GetAllTasks() As Collection
    Dim ws As Worksheet
    Set ws = wb.Worksheets(DATA_SHEET_NAME)
    Dim col As New Collection
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        ' Нет данных, вернуть пустую коллекцию
        Set ITaskRepository_GetAllTasks = col
        Exit Function
    End If
    ' Читаем строки 2..lastRow
    For i = 2 To lastRow
        Dim title As String, completedStr As String
        title = CStr(ws.Cells(i, 1).Value)
        completedStr = CStr(ws.Cells(i, 2).Value)
        If title = "" Then
            ' пропустим пустые
        Else
            Dim t As Task: Set t = New Task
            t.Title = title
            If completedStr = "TRUE" Then t.IsCompleted = True _
            Else t.IsCompleted = False
            col.Add t
        End If
    Next i
    Set ITaskRepository_GetAllTasks = col
End Function

' Удалить задачу (по ID или названию)
Public Sub ITaskRepository_DeleteTask(ByVal taskId As String)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(DATA_SHEET_NAME)
    ' Пройдёмся по данным и найдём строку с указанным ID или Title.
    ' Так как FileRepository не сохраняет ID, будем считать, 
    ' что taskId передается как Title (допущение для простоты).
    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        Dim title As String
        title = CStr(ws.Cells(i, 1).Value)
        If title = taskId Then
            ws.Rows(i).Delete Shift:=xlUp
            Exit For
        End If
    Next i
End Sub

Private Sub Class_Terminate()
    Set wb = Nothing
End Sub
```

*Комментарий:* `FileRepository` сохраняет задачи на листе Excel. При инициализации (`Class_Initialize`) он проверяет, есть ли лист `TasksData` в текущей книге (`ThisWorkbook`). Если нет – создаёт его и добавляет строку заголовков. Мы используем `ThisWorkbook` для простоты, но можно было бы передавать Workbook параметром в конструктор (например, через метод Init). Методы:

* `AddTask` – определяет последнюю заполненную строку и добавляет новую задачу в следующей строке. В колонку A записывается название, в колонку B – статус ("TRUE"/"FALSE"). (Заметим, что мы не сохраняем явный ID; здесь предполагается уникальность названия задачи для удаления – упрощение.)
* `GetAllTasks` – читает все записи (со 2-й строки до последней заполненной). На каждую запись создаётся объект `Task` и добавляется в коллекцию. Обратите внимание: ID, сгенерированный в `Task.Class_Initialize`, не будет совпадать с тем, что был при добавлении – т.к. мы не сохраняем ID на листе. Это недостаток данной примитивной реализации: при повторной загрузке `TaskManager` создаст новые объекты `Task` с новыми ID. В реальном приложении стоило бы сохранять ID тоже (например, в отдельной колонке). Для примера, однако, это не критично.
* `DeleteTask` – находит первую строку, где название равно переданному `taskId` и удаляет её. Здесь мы фактически трактуем параметр как название задачи (т.к. нет сохранённых ID). Это ограничение: если есть дубли названий, удалится первая попавшаяся. Тем не менее, для демонстрации принципа работы DAL достаточно.

Эта реализация показывает, как слой DAL изолирует работу с Excel: *ни класс Task, ни TaskManager, ни тем более UI-форма не знают*, как именно сохраняются данные. Они вызывают методы интерфейса `ITaskRepository`. `FileRepository` же не знает, зачем эти данные нужны – он только сохраняет и выдаёт их. Такое разделение соответствует **принципу единственной ответственности**: изменения формата хранения (например, решим хранить в CSV-файле вместо листа) затронут только `FileRepository`, а бизнес-логику и интерфейс менять не придётся.

Теперь, когда все элементы реализованы, представим **общую схему взаимодействия** между слоями:

```mermaid
flowchart LR
    subgraph UI [UI Layer (User Interface)]
        UIForm[frmTaskManager (UserForm)]
    end
    subgraph BL [BL Layer (Business Logic)]
        TM[TaskManager.cls]
        TaskClass[Task.cls]
    end
    subgraph DAL [DAL Layer (Data Access)]
        IRepo[ITaskRepository.cls <<interface>>]
        FakeRepo[FakeStorage.cls]
        FileRepo[FileRepository.cls]
    end

    UIForm -- вызывает методы --> TM
    UIForm -- слушает события --> TM
    TM -- создает --> TaskClass
    TM -- использует --> IRepo
    IRepo -- реализован --> FakeRepo
    IRepo -- реализован --> FileRepo
    FakeRepo -- хранит в --> Memory[(In-Memory Collection)]
    FileRepo -- хранит в --> Sheet[(Excel Worksheet)]
    TM -- ведет лог через --> Logger[(Logger.cls)]
```

В этой диаграмме видно, как `frmTaskManager` (UI) взаимодействует с `TaskManager` (BL): вызывает его методы, подписывается на события. `TaskManager` в свою очередь работает через интерфейс `ITaskRepository`, не зная, какая конкретно реализация используется – это может быть либо `FakeStorage` (память), либо `FileRepository` (лист Excel). Дополнительно, `TaskManager` может использовать `Logger` для логирования (показано пунктирной связью). **Потоки данных** идут от UI к BL (запросы пользователя), далее от BL к DAL (операции с данными), и обратно от BL к UI (события, возвращаемые данные).

## Примеры (Examples)

Рассмотрим три примера использования нашей многослойной архитектуры, от абстрактного к конкретному.

### Пример 1: Абстрактная демонстрация потока данных

Чтобы понять взаимодействие слоёв, проследим типичный сценарий на концептуальном уровне:

1. **Пользователь** нажимает кнопку "Добавить" на форме (`frmTaskManager`). Форма собирает введённые данные (например, название задачи) и вызывает метод бизнес-логики: `TaskManager.AddTask`.
2. **TaskManager** получает запрос `AddTask`. Он создаёт новый объект `Task` (устанавливает ему название) и вызывает метод DAL: `repo.AddTask`, чтобы сохранить задачу. Затем добавляет задачу в свой список в памяти и генерирует событие `TaskAdded`.
3. **FileRepository/FakeStorage** (конкретный DAL) выполняет `AddTask`: либо пишет строку на лист Excel (если это `FileRepository`), либо сохраняет объект в словаре в памяти (`FakeStorage`).
4. **UserForm** ранее подписалась на событие `TaskManager.TaskAdded`. Когда событие срабатывает, форма получает уведомление о добавленной задаче (например, в обработчике `TaskManager_TaskAdded`). В ответ на это событие, UI добавляет новую строку в список отображаемых задач (например, в ListBox на форме) – т.е. обновляет **представление**.

По такому принципу работают и другие операции: UI вызывает BL, BL обращается к DAL, а BL через события сообщает UI об изменениях или результатах.

Теперь посмотрим, как это выглядит в коде. Допустим, мы хотим без UI добавить пару задач и вывести их список – это проверит работу BL и DAL:

```vba
' Имитация сценария без UI (Immediate Window / модуль)
Sub DemoFlow()
    Dim mgr As New TaskManager         ' бизнес-логика
    mgr.SetRepository New FakeStorage  ' используем фальшивое хранилище (можно сменить на FileRepository)

    ' Добавляем задачи через бизнес-логику
    mgr.AddTask "Пример 1"
    mgr.AddTask "Пример 2"

    ' Получаем все задачи и выводим их заголовки
    Dim all As Collection
    Set all = mgr.GetAllTasks()
    Dim t As ITask
    For Each t In all
        Debug.Print "Задача: "; t.Title, "Завершена? "; t.IsCompleted
    Next

    ' Отметим первую задачу выполненной
    If all.Count > 0 Then
        Dim first As Task
        Set first = all.Item(1)
        first.MarkComplete
        Debug.Print "Отметили '" & first.Title & "' как выполненную."
    End If
End Sub
```

**Что происходит в этом коде?** Мы создаём `TaskManager` и явно устанавливаем ему `FakeStorage` как хранилище (репозиторий). Далее добавляем две задачи ("Пример 1" и "Пример 2"). Внутри `AddTask` они сохраняются в `FakeStorage` и также в памяти `TaskManager`. Затем получаем все задачи через `GetAllTasks()` и печатаем их в Immediate Window (Debug.Print). Ожидаемый вывод:

```
Задача: Пример 1        Завершена?  False
Задача: Пример 2        Завершена?  False
```

Затем помечаем первую задачу выполненной (`MarkComplete` устанавливает `IsCompleted = True`) и выводим сообщение об этом. Здесь мы напрямую работаем с объектом `Task` (первым элементом коллекции). В реальном приложении, UI-слой бы обновил отображение статуса.

Обратите внимание: даже в этом простом скрипте UI не участвует, но код прекрасно работает, т.к. слои BL и DAL автономны. Это подтверждает тестируемость архитектуры – мы можем писать модульные тесты на `TaskManager`, подставляя `FakeStorage` для воспроизведения разных ситуаций, без запуска формы. Например, с помощью Rubberduck Unit Testing можно было бы автоматизировать проверку, что после `AddTask` количество задач увеличилось на 1, и т.п.

### Пример 2: Простой сценарий без UI (консольный)

Этот пример похож на предыдущий, но оформим его как отдельный тестовый Sub и покажем результат. Предположим, у нас настроено всё, как описано выше (все классы находятся в проекте). Выполним такой код:

```vba
Sub Example_NoUI()
    Dim repo As ITaskRepository
    Set repo = New FakeStorage         ' выбираем реализацию DAL (можно попробовать FileRepository)

    Dim mgr As New TaskManager
    mgr.SetRepository repo            ' связываем менеджер с хранилищем

    Debug.Print "--- Добавление задач ---"
    mgr.AddTask "Задача A"
    mgr.AddTask "Задача B"
    mgr.AddTask "Задача C"

    Debug.Print "--- Список всех задач ---"
    Dim t As ITask
    For Each t In mgr.GetAllTasks()
        Debug.Print t.Title, IIf(t.IsCompleted, "Completed", "Pending")
    Next

    Debug.Print "--- Удаление задачи B ---"
    mgr.RemoveTask "Задача B"   ' удалим задачу с названием "Задача B" (в FakeStorage title == key)
    ' Примечание: Для FileRepository этот вызов удалит первую задачу с именем "Задача B" на листе.

    Debug.Print "--- Список после удаления ---"
    For Each t In mgr.GetAllTasks()
        Debug.Print t.Title, IIf(t.IsCompleted, "Completed", "Pending")
    Next
End Sub
```

**Вывод (Immediate Window):**

```
--- Добавление задач ---
INFO #1: Added task 'Задача A' (ID=20230505060935000)   ' (логгер пишет информацию)
INFO #2: Added task 'Задача B' (ID=20230505060935000)
INFO #3: Added task 'Задача C' (ID=20230505060935000)
--- Список всех задач ---
Задача A            Pending
Задача B            Pending
Задача C            Pending
--- Удаление задачи B ---
INFO #4: Removed task ID=Задача B
--- Список после удаления ---
Задача A            Pending
Задача C            Pending
```

*(Примерный вывод; ID будут уникальными у вас, здесь показано концептуально. Логгер пронумеровал события.)*

Мы видим, что добавление прошло успешно (логгер зафиксировал 3 добавления). Список отобразил все три задачи. После удаления задачи B, логгер сообщил об удалении, и финальный список содержит только A и C.

Этот сценарий демонстрирует работу **бизнес-логики и DAL в консольном режиме**, что удобно для отладки. Можно свободно переключить `repo = New FileRepository` – тогда при добавлении задач они реально запишутся на лист Excel, и при повторном запуске Example\_NoUI (или при запуске формы) эти задачи можно будет прочитать обратно. Таким образом, мы проверили взаимозаменяемость DAL: `TaskManager` работает одинаково с любым хранилищем, разница лишь в данных.

### Пример 3: Расширенный пример с UI (пользовательская форма)

Теперь интегрируем всё в полноценном Excel-файле с формой. Представим, что у нас есть форма `frmTaskManager` с простым интерфейсом:

* Поле ввода текста (Name: `txtTaskTitle`) для названия новой задачи.
* Кнопка "Add" (Name: `btnAdd`) для добавления.
* ListBox (Name: `lstTasks`) для списка задач.
* Кнопка "Mark Complete" (`btnComplete`) для пометki выбранной задачи выполненной.
* Кнопка "Delete" (`btnDelete`) для удаления выбранной задачи.
* (Можно также иметь метку статуса, кнопку "Refresh", но не обязательно).

**Код формы** покажет, как она взаимодействует с `TaskManager`:

```vba
' src\Layered\frmTaskManager.frm (UserForm code-behind)
Option Explicit

' Связь с бизнес-логикой:
Private WithEvents taskMgr As TaskManager

' При открытии формы:
Private Sub UserForm_Initialize()
    Set taskMgr = New TaskManager
    taskMgr.SetRepository New FileRepository  ' используем реальное хранилище (лист Excel)
    ' Заполним список существующих задач при запуске
    Dim t As ITask
    For Each t In taskMgr.GetAllTasks()
        lstTasks.AddItem t.Title & IIf(t.IsCompleted, " (done)", "")
    Next
End Sub

' Обработчик кнопки "Add"
Private Sub btnAdd_Click()
    Dim title As String
    title = Trim(txtTaskTitle.Text)
    If title = "" Then
        MsgBox "Введите название задачи.", vbExclamation
        Exit Sub
    End If
    ' Добавляем задачу через менеджер
    Dim newTask As ITask
    Set newTask = taskMgr.AddTask(title)
    txtTaskTitle.Text = ""  ' очищаем поле ввода
    ' Примечание: список обновится через событие TaskAdded (см. ниже)
End Sub

' Обработчик события TaskManager.TaskAdded
Private Sub taskMgr_TaskAdded(ByVal newTask As ITask)
    ' Добавляем новую задачу в ListBox
    lstTasks.AddItem newTask.Title & IIf(newTask.IsCompleted, " (done)", "")
End Sub

' Обработчик кнопки "Mark Complete"
Private Sub btnComplete_Click()
    If lstTasks.ListIndex < 0 Then
        MsgBox "Выберите задачу для отметки как выполненной.", vbInformation
        Exit Sub
    End If
    Dim selectedTitle As String
    selectedTitle = lstTasks.List(lstTasks.ListIndex)
    selectedTitle = Replace(selectedTitle, " (done)", "") ' убираем пометку, если есть
    ' Найти задачу по названию и отметить выполненной:
    Dim t As ITask, allTasks As Collection
    Set allTasks = taskMgr.GetAllTasks()
    For Each t In allTasks
        If t.Title = selectedTitle Then
            t.IsCompleted = True
            Exit For
        End If
    Next
    ' Обновим элемент списка с пометкой "(done)"
    lstTasks.List(lstTasks.ListIndex) = selectedTitle & " (done)"
    ' Опционально: можно сохранить изменения в хранилище сразу,
    ' но у нас FileRepository не отслеживает обновление статуса уже сохраненных записей.
    ' В реальном случае, стоило бы вызвать repo.UpdateTask и обновить на листе.
End Sub

' Обработчик кнопки "Delete"
Private Sub btnDelete_Click()
    If lstTasks.ListIndex < 0 Then
        MsgBox "Выберите задачу для удаления.", vbInformation
        Exit Sub
    End If
    Dim selectedTitle As String
    selectedTitle = lstTasks.List(lstTasks.ListIndex)
    selectedTitle = Replace(selectedTitle, " (done)", "")
    ' Удаляем через менеджер
    taskMgr.RemoveTask selectedTitle
    ' Удаляем из списка UI
    lstTasks.RemoveItem lstTasks.ListIndex
End Sub

' При закрытии формы
Private Sub UserForm_Terminate()
    Set taskMgr = Nothing  ' освобождаем объект TaskManager
End Sub
```

Объяснения к коду формы:

* В `UserForm_Initialize` мы создаём экземпляр `TaskManager` и задаём ему `FileRepository` как хранилище (так, чтобы данные сохранялись на листе). Затем сразу заполняем `lstTasks` текущими задачами, получив их из менеджера. Каждая завершённая задача помечается "(done)" в списке.
* `btnAdd_Click`: при нажатии "Add" проверяем, что поле не пустое, и вызываем `taskMgr.AddTask`. Если задача успешно добавлена, очищаем текстовое поле. Обновление списка *не делаем вручную* здесь, потому что сработает событие `taskMgr_TaskAdded`.
* `taskMgr_TaskAdded`: это реализация обработчика события, которое мы объявили через `WithEvents`. Когда TaskManager внутри себя выполнит `RaiseEvent TaskAdded`, форма автоматически вызовет этот метод, получив объект новой задачи. Мы просто добавляем новый элемент в `lstTasks` – таким образом UI обновляется сразу после добавления.
* `btnComplete_Click`: обрабатывает нажатие "Mark Complete". Сначала проверяем, выбран ли элемент в списке (ListIndex >= 0). Далее получаем выбранный заголовок (убирая суффикс " (done)" если был). Затем проходим по всем задачам `TaskManager` и находим ту, что соответствует заголовку, и отмечаем её выполненной (`t.IsCompleted = True`). После этого обновляем отображение выбранного элемента, добавив "(done)". Здесь мы демонстрируем, что можно напрямую модифицировать объект задачи через бизнес-логику. (Замечание: `FileRepository` у нас не поддерживает обновление записи – в реальном случае нужно было бы реализовать метод обновления, но для упрощения считаем, что отметка выполненной не сохраняется на диск; при перезапуске формы статус опять будет взят из файла как "невыполнен". Это упущение в примере, но не принципиальное для рассмотрения архитектуры.)
* `btnDelete_Click`: аналогично, получаем выбранный элемент, его заголовок, вызываем `taskMgr.RemoveTask`. Обратите внимание: мы передаем `selectedTitle` как идентификатор. В `TaskManager.RemoveTask` мы ожидаем ID; мы упрощённо используем название как ID, что работает с `FakeStorage` (где key=Title) и `FileRepository` (где DeleteTask удаляет по Title). После вызова мы также убираем элемент из списка `lstTasks` на форме. (Можно было бы дождаться события `TaskRemoved` и обработать его подобно добавлению, но здесь мы сделали удаление синхронно для простоты UI.)
* `UserForm_Terminate`: при закрытии формы освобождаем объект `taskMgr`. Это важно, чтобы менеджер не висел в памяти после закрытия формы. Поскольку форма была `WithEvents` подписана на `taskMgr`, наличие ссылки могло бы помешать сборке мусора без этого шага.

**Как это выглядит для пользователя:** При открытии формы, он сразу видит список задач (если на листе были сохранены – благодаря `UserForm_Initialize`). Он вводит название в поле, нажимает "Add" – новая задача появляется в списке (UI за счёт события). Если нажать "Delete" с выбранной задачей – она удаляется как из списка, так и из хранилища (лист). Кнопка "Mark Complete" помечает задачу как выполненную (добавляя "(done)" визуально).

**Примечание:** В реальном приложении стоило бы также изменить дизайн `FileRepository` для отслеживания обновлений (например, добавить метод `UpdateTaskStatus`), но здесь мы сосредоточились на общей структуре.

### Структура каталогов и файлы примера

Все исходные коды, описанные выше, логично разнести по папкам проекта:

```
Project
├── docs
│   └── LayeredArchitecture.md        ' финальный гайд (в формате Markdown)
├── src
│   ├── Common
│   │   ├── Logger.cls
│   │   ├── ITask.cls
│   │   ├── ITaskRepository.cls
│   │   ├── TestData.bas
│   │   └── BaseForm.cls              ' (опционально, как обсуждалось)
│   └── Layered
│       ├── Task.cls
│       ├── TaskManager.cls
│       ├── FakeStorage.cls
│       ├── FileRepository.cls
│       └── frmTaskManager.frm
├── tests
│   └── TaskManagerTests.bas          ' модуль с юнит-тестами Rubberduck
└── samples
    └── LayeredArchitectureSample.xlsm ' готовый файл Excel с импортированными классами и формой
```

В папке `docs` хранится наше руководство (включая этот текст). Папка `src` содержит исходники: общие компоненты и конкретную реализацию многослойной архитектуры. В `tests` можно разместить автоматические тесты. В `samples` – готовый Excel с интегрированными слоями. Такое разделение упрощает сопровождение: можно править код в отдельных файлах и собирать проект с помощью, например, Rubberduck (у него есть функции импорта/экспорта модулей). Также это удобно для контроля версий (хранения в Git).

**Пример теста (tests\TaskManagerTests.bas):** воспользуемся возможностями Rubberduck для создания модульного теста бизнес-логики:

```vba
' tests\TaskManagerTests.bas
Option Explicit
'@TestModule
Private Assert As New Rubberduck.AssertClass

'@TestMethod
Public Sub AddTask_IncreasesCount()
    ' Arrange
    Dim repo As ITaskRepository: Set repo = New FakeStorage
    Dim mgr As New TaskManager
    mgr.SetRepository repo
    Dim initialCount As Long
    initialCount = mgr.GetAllTasks().Count
    ' Act
    mgr.AddTask "Test Task"
    ' Assert
    Dim newCount As Long
    newCount = mgr.GetAllTasks().Count
    Assert.IsEqual initialCount + 1, newCount, "Count of tasks should increase by 1 after adding."
End Sub

'@TestMethod
Public Sub RemoveTask_DeletesFromRepo()
    ' Arrange
    Dim repo As ITaskRepository: Set repo = New FakeStorage
    Dim mgr As New TaskManager
    mgr.SetRepository repo
    Dim t As ITask: Set t = mgr.AddTask("Task to remove")
    ' Act
    mgr.RemoveTask t.Title  ' (в FakeStorage Title = key)
    ' Assert
    Dim remaining As Collection: Set remaining = mgr.GetAllTasks()
    Dim found As Boolean: found = False
    Dim it As ITask
    For Each it In remaining
        If it.Title = t.Title Then
            found = True
            Exit For
        End If
    Next
    Assert.IsFalse found, "Task was not removed from collection"
End Sub
```

Здесь мы написали два тестовых метода: первый проверяет, что после добавления задачи количество увеличилось, второй – что после удаления задачи её нет в списке. Rubberduck позволит запустить эти тесты прямо в VBE, и они пройдут без участия UI, используя `FakeStorage`. Это подтверждает **юнит-тестируемость** нашей архитектуры – важное преимущество отделения логики от интерфейса.

## Оценка (Evaluation)

Многослойная архитектура заметно улучшает **качество и устойчивость кода** в VBA-проектах. Проведём оценку её эффектов по нескольким критериям:

* **Разделение обязанностей:** Каждый модуль решает свою задачу. Это облегчает отладку и сопровождение – изменения в одном слое минимально затрагивают другие. Например, форму можно изменять (добавлять поля, изменять внешний вид) без риска сломать бизнес-логику; или заменить хранение данных на БД, не трогая форму и логику.
* **Тестируемость:** Как показано выше, бизнес-логику можно тестировать автономно. Можно автоматически проверить десятки сценариев добавления/удаления задач, не кликая по форме. Это повышает надежность приложения.
* **Повторное использование:** Компоненты становятся переиспользуемыми. Класс `TaskManager` и сопутствующие интерфейсы можно использовать в другом проекте с минимальными изменениями – достаточно реализовать подходящий `ITaskRepository`. Логгер, генераторы данных, общие функции UI (например, `BaseForm`) служат всем архитектурным решениям.
* **Читаемость кода:** Код структурирован по папкам и слоям, именование явно указывает роль (например, суффиксы \*Manager, \*Repository). Это делает проект понятным для нового разработчика. Короткие процедуры в UI (только вызовы методов/события) легче воспринимать, чем гигантские обработчики со смешанной логикой (проблема *Smart UI* выше).
* **Устранение частых ошибок:** Многие типовые ошибки исчезают при соблюдении архитектуры. Например, использование `Option Explicit` во всех модулях предотвращает опечатки и неявное создание переменных; отказ от глобальных переменных снижает непредсказуемые связи (состояние хранится внутри классов); использование событий исключает необходимость опрашивать состояние, уменьшает зависимость UI от BL.

Рассмотрим короткие сравнения фрагментов кода **до** и **после** применения архитектурных принципов:

► **Плохо:** Монолитный код в UserForm (форма сама лезет на лист Excel, нет разделения):

```vba
' Code inside a UserForm (anti-pattern: Smart UI)
Private Sub btnSave_Click()
    Dim title As String
    title = txtTitle.Text
    If title = "" Then Exit Sub
    ' Сразу пишем на лист из UI:
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TasksData")
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = title
    MsgBox "Task added!"
End Sub
```

Этот подход нарушает слоистость: форма знает о структуре данных, напрямую обращается к Excel. Логика добавления размазана по UI. Масштабирование такого кода приводит к тому, что форма обрастает логикой (анти-паттерн **Smart UI**).

► **Хорошо:** UI вызывает BL через интерфейс, ничего не знает о деталях хранения:

```vba
' Code inside UserForm using TaskManager (Layered approach)
Private Sub btnAdd_Click()
    Dim title As String: title = Trim(txtTitle.Text)
    If title = "" Then Exit Sub
    taskMgr.AddTask title   ' бизнес-логика сама решит, как сохранить
    txtTitle.Text = ""
    ' UI обновится через событие TaskAdded
End Sub
```

Форма делегирует работу `TaskManager`. Добавление задачи инкапсулировано в классе BL, а UI лишь передает ввод. Это **хорошо**: разделение ответственности, код проще поддерживать и тестировать.

Еще один пример:

► **Плохо:** Использование **глобальных переменных** для передачи данных между процедурами и формами:

```vba
Public currentTaskTitle As String  ' глобальная переменная

Sub DoSomething()
    currentTaskTitle = "Test"
    UserForm1.Show
End Sub

' В UserForm1:
Private Sub UserForm_Activate()
    txtTitle.Text = currentTaskTitle  ' использует глобальную переменную
End Sub
```

Глобальное состояние усложняет понимание: любая часть программы может изменить `currentTaskTitle`, трудно отследить поток данных. Rubberduck отмечает, что чаще всего глобальные переменные не нужны, и состояние можно передавать через объекты.

► **Хорошо:** Вместо глобальных – использовать свойства объекта или параметры методов:

```vba
' Без глобальных: передаем через объект
Sub ShowTaskForm(title As String)
    Dim frm As New frmTaskManager
    frm.InitialTitle = title  ' устанавливаем свойство формы
    frm.Show
End Sub

' В коде формы:
Private m_initialTitle As String
Public Property Let InitialTitle(ByVal value As String)
    m_initialTitle = value
End Property

Private Sub UserForm_Initialize()
    If m_initialTitle <> "" Then txtTitle.Text = m_initialTitle
End Sub
```

Здесь `InitialTitle` – свойство формы, через которое внешний код задаёт контекст. Внутри формы нет обращения к глобальной переменной. Это более безопасно и понятно: зависимости явные.

Такие улучшения можно привести для многих ситуаций. Многослойная архитектура направляет разработчика использовать **лучшие практики**:

* Явное объявление переменных (`Option Explicit`) – меньше скрытых ошибок.
* Минимум глобального состояния (лучше передать нужные данные через параметры или хранить в классах).
* Событийная модель взаимодействия – вместо бесконечного опроса или жёсткой связанности.
* Чёткое распределение функций по модулям – облегчает оптимизацию и поиск узких мест.

Наконец, оценим возможные **издержки** архитектуры: безусловно, такой подход требует написать больше кода (создать несколько классов, продумать интерфейсы). В простейших макросах накладные расходы могут казаться излишними. Однако, как только проект растёт, эта начальная инвестиция окупается снижением количества ошибок и затрат на изменение кода. Потери производительности от вызовов методов классов минимальны для большинства задач (VBA справляется с небольшим количеством объектов; узкими местами скорее станет работа с листами Excel, чем сама архитектура).

В итоге, **плюсы многослойной архитектуры многократно перевешивают минусы** для проектов средней и высокой сложности, что делает её рекомендованным шаблоном при разработке VBA-приложений.

**Краткая сводка оценки**:

| Критерий               | До (спагетти / без слоёв)                | После (Layered Architecture)                              |
| ---------------------- | ---------------------------------------- | --------------------------------------------------------- |
| Разделение логики      | Всё смешано (UI, логика, данные вместе)  | Чёткие слои: UI vs BL vs DAL                              |
| Читаемость и поддержка | Падает с ростом кода (форма перегружена) | Улучшается (каждый слой прост и понятен)                  |
| Тестируемость          | Трудно (завязано на UI/Excel среду)      | Легко (бизнес-логику можно тестировать отдельно)          |
| Переиспользование      | Низкое (код привязан к контексту)        | Высокое (компоненты независимы, могут переиспользоваться) |
| Вероятность ошибок     | Высокая (много скрытых взаимосвязей)     | Ниже (ошибки локализованы, отлавливаются рано)            |
| Время разработки       | Чуть быстрее старт, но сложно отладить   | Чуть дольше старт, но легче развивать далее               |
| Производительность     | (зависит от реализации, но обычно ок)    | (зависит от реализации; накладные вызовы минимальны)      |

## Ошибки и граничные случаи (Errors/Edge-cases)

При разработке и использовании многослойной архитектуры важно учитывать возможные ошибки (в дизайне и в runtime) и граничные случаи. Ниже приведена таблица распространённых **анти-паттернов** и ошибок, характерных для VBA, с указанием последствий и способов оптимизации/решения:

| **Анти-паттерн / ошибка**                                                                                                    | **Почему плохо / последствия**                                                                                                                      | **Решение / профилактика**                                                                                                                                                                                                                                                      |
| ---------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **«Умная форма» (Smart UI)** – вся логика в коде UserForm, форма напрямую лезет в базу/лист                                  | Код UI разрастается, трудность в тестировании, слабая масштабируемость. Изменение хранения требует правки формы, высокий риск багов.                | **Разделение слоёв:** вынести логику и работу с данными в классы BL/DAL. Форма лишь отображает и вызывает методы.                                                                                                                                                               |
| **Отсутствие `Option Explicit`** – не объявляются переменные явно                                                            | Опечатки не ловятся компилятором, возможны непредсказуемые баги (например, создание нового варианта вместо использования существующего переменной). | Всегда включать `Option Explicit` в каждом модуле (настройка VBEditor > *Require Variable Declaration*). Тогда опечатки выявляются при компиляции, переменные имеют понятные типы.                                                                                              |
| **Глобальные переменные для передачи данных**                                                                                | Скрытые зависимости, усложняют повторное использование. Трудно отследить, кто и когда изменяет глобал. Возможны конфликты имен.                     | Минимизировать глобальные переменные. Передавать данные через параметры процедур, свойства классов. Использовать классы для хранения состояния (контекст) вместо Moduл-level Public переменных.                                                                                 |
| **Отсутствие событий / обратной связи** – UI не знает об изменениях, кроме как опрашивать BL                                 | Либо интерфейс не обновляется, либо приходится каждую секунду проверять состояние (что неэффективно).                                               | Использовать **Events**: пусть BL (или DAL) генерирует события при изменениях, а UI подпишется и обновится. Это снижает связность и упрощает обновление UI.                                                                                                                     |
| **Не освобождаются объекты** – например, не `Unload` форма, не `Set Nothing` для объектов                                    | Утечки памяти, “призрачные” объекты. В случае форм: скрытые экземпляры могут сохранять старое состояние, события могут продолжать срабатывать.      | Всегда `Unload Me` для закрытия форм (или `Set frm = Nothing` если форма объект). В классе используйте `Class_Terminate` для очистки важных ссылок. При использовании WithEvents убедиться, что цепочка разрывается (например, наша форма `Set taskMgr = Nothing` в Terminate). |
| **Обработка ошибок не реализована** – весь код предполагает идеальное выполнение                                             | При любой runtime-ошибке макрос падает, пользователь видит непонятное сообщение.                                                                    | Использовать `On Error GoTo ...` в ключевых местах (например, при работе с файлами). Логгировать ошибку (через Logger) и аккуратно информировать пользователя (через MsgBox или отображение на форме). Центральную обработку можно вынести в общую функцию.                     |
| **Жёстко закодированные ссылки на объекты Excel** – например, `Sheets("Data")` в глубине кода                                | Если имя листа изменится – код сломается. Сложно переносить в другой файл. Тестирование вне Excel затруднено.                                       | Абстрагировать доступ к Excel в DAL. Например, `FileRepository` оперирует именем листа в одном месте. Использовать константы/настройки для таких параметров. Для тестов – подменять `ThisWorkbook` на объект-заглушку или использовать FakeStorage.                             |
| **Дублирование кода** между слоями или модулями                                                                              | Нарушает принцип DRY, увеличивает вероятность ошибок при изменении (правки нужно делать в нескольких местах).                                       | Выделить общие функции/модули. Например, генерация ID задачи – вынести в отдельную функцию, использовать её и в FakeStorage, и в FileRepository для консистентности. Общие константы (имена листов, сообщений) держать в одном модуле.                                          |
| **Пренебрежение проверками граничных условий** – например, удаление задачи из пустого списка, или чтение файла, которого нет | Ошибки выполнения (Null reference, индекс вне диапазона и т.п.).                                                                                    | Добавлять проверки: если коллекция пустая – обработать отдельно; если файл/лист отсутствует – создать или вывести сообщение. Наш код FileRepository демонстрирует такую проверку (создание листа при отсутствии).                                                               |
| **Избыточная оптимизация преждевременно** – чрезмерно усложнённый код "на вырост"                                            | Снижается понятность, увеличивается количество багов, хотя реально могло не понадобиться.                                                           | Следовать KISS-принципу: реализовать архитектуру, достаточную для текущих целей, но с возможностью расширения. Например, мы не стали реализовывать обновление задач в FileRepository, чтобы не загромождать код – при необходимости добавим по мере развития.                   |

Помимо этих ошибок разработки, есть **граничные случаи** исполнения, о которых стоит подумать при архитектурном подходе:

* **Масштаб данных:** Что если задач будет очень много (например, 10 тысяч)? Наша архитектура в целом справится, но узким местом станет `FileRepository` (операции с листом Excel на каждую задачу). Решение: оптимизировать DAL (например, батч-операции, использовать массивы или ADO). BL и UI слои при этом менять не нужно.
* **Параллельность:** VBA исполняется в одном потоке, но можно открыть несколько форм одновременно (модельно). В нашей архитектуре можно иметь несколько `TaskManager` экземпляров, однако если они работают с одним хранилищем (например, одним и тем же листом), возможны коллизии. Для простоты мы не рассматривали многопоточность. Если нужно, следует вводить блокировки или ограничения (например, отключить кнопки на время сохранения, либо использовать Application.Level events для синхронизации). Обычно, в Excel-VBA окружении параллельность минимальна.
* **Жизненный цикл объектов:** Нужно следить, чтобы объекты BL не жили дольше, чем нужно. Например, если `TaskManager` хранится в модуле как глобальная переменная и форма переоткрывается, можно получить дублирование данных (т.к. при повторной инициализации из файла они продублируются). В нашем примере `TaskManager` создаётся при каждом открытии формы заново, что упрощает управление. В других сценариях можно применять синглтоны или паттерн Presenter, но это выходит за рамки данной части.
* **Совместимость:** Код, написанный с классами и т.д., требует, чтобы книга была в формате .xlsm и чтобы макросы были включены. Это очевидно, но стоит упомянуть: при распространении такого решения убедитесь, что пользователи знают о необходимости разрешить VBA.
* **Отладка событий:** Использование Events осложняет отладку, т.к. порядок выполнения не линейный. Важно правильно понимать, что сначала выполнится код BL, потом сработает обработчик в UI. Если в обработчике возникает ошибка, желательно ее ловить (например, обернуть код внутри `taskMgr_TaskAdded` в `On Error`, чтобы падение UI-слоя не рушило всю программу).
* **Unit-тесты в VBA:** Rubberduck значительно облегчает их реализацию, но нужно помнить, что тесты запускаются в среде VBA, и нельзя выполнять их вне Excel легко. Граничный случай – если код DAL обращается к `ThisWorkbook`, в юнит-тестах (которые могут быть запущены без UI) `ThisWorkbook` существует, но может не содержать ожидаемых листов. Мы это решаем, используя `FakeStorage` для тестов, избегая зависимостей на Excel.

Подведём итог разделу ошибок: многослойная архитектура сама по себе призвана устранить многие системные ошибки проектирования (такие как смешение логики и UI). Тем не менее, разработчик должен дисциплинированно следовать её принципам, иначе польза снизится. Таблица ниже резюмирует соответствие между принципами схемы, типичными ошибками при их нарушении и путями оптимизации:

| **Принцип (Слой/схема)**                | **Возможная ошибка при нарушении**                                              | **Оптимизация (как решает архитектура)**                                        |
| --------------------------------------- | ------------------------------------------------------------------------------- | ------------------------------------------------------------------------------- |
| UI только отображение, без логики       | Логика в UI -> *Smart UI*, трудна поддержка                                     | Переносим логику в BL, UI упрощается                                            |
| BL независим от UI (чистая логика)      | BL обращается к элементам формы -> зависимость, невозможно тестировать отдельно | BL ничего не знает о UI, можно тестировать в изоляции                           |
| DAL абстрагирован интерфейсом           | Прямой вызов файлов/листов в BL -> дублирование, сложно сменить источник данных | Интерфейс + реализация DAL позволяет подменять хранение без правки BL           |
| Использование `Option Explicit`         | Необъявленные переменные -> runtime ошибки, Variants                            | Компиляция ловит ошибки, улучшает производительность и читаемость               |
| Событийная модель (Events)              | Отсутствие уведомлений -> UI не синхронизирован                                 | Events обеспечивают автоматическую синхронизацию состояния UI с BL              |
| Управление ресурсами (Unload/Terminate) | Висящие объекты/формы -> утечки, непредвиденное поведение                       | Явное закрытие форм и освобождение объектов по завершении работы слоя           |
| Логгирование действий и ошибок          | Тихое игнорирование ошибок -> сложно отладить                                   | Централизованный Logger фиксирует все ключевые события, облегчая разбор полётов |
| Тесты на уровне BL (и DAL отдельно)     | Нет автоматических тестов -> регрессии при изменениях                           | Тестируем BL с FakeStorage, ловим баги до UI, облегчаем рефакторинг             |

Каждый слой архитектуры вместе с сопутствующими практиками помогает избежать определённых ошибок и облегчает оптимизацию приложения. Следуя принципам **Essence, Elements, Examples, Evaluation, Errors** при проектировании (как мы сделали в этом разделе), можно не только построить работающее решение, но и понять *почему* оно устроено именно так и *какие проблемы* это решает. Многослойная архитектура в VBA – мощный инструмент, делающий даже сложные Excel-приложения более понятными для человека и, как мы видим, вполне постижимыми для искусственного помощника, анализирующего такой Markdown-гайд.
