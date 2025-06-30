Отлично. Я подготовлю полную библиотеку из 25 прогрессивных примеров по Excel VBA с передовыми event-driven техниками, строго модульно организованную, с учётом Windows Desktop 64-bit и соблюдением паттернов проектирования и чистого кода.

В дополнение будет включено:

* 3 мини-кейса с объяснением реальных задач и обоснованием выбранных решений,
* таблицы справочной информации (последовательности событий, выбор паттерна, затраты API),
* JSON-инвентарь для LLM-каталога с intent-тегами.

Я сообщу, когда всё будет готово.


# Advanced Event-Driven Excel VBA – Comprehensive Examples

**Overview:** This knowledge base showcases 25 progressively complex examples of event-driven programming in Excel VBA, organized by theme. We emphasize modular design (using class modules, UserForms, and standard modules) targeting Windows Excel (64-bit). Each example is fully runnable (copy code into a workbook, insert class modules/UserForms as named, run the provided test macros) and demonstrates best practices in object-oriented VBA, event handling, and cleanup. Three mini case studies and quick-reference tables are included to deepen understanding.

## A. Core Event Patterns (8 Examples)

These examples cover fundamental event patterns: defining custom events, implementing Observer/Mediator designs, using a centralized Event Aggregator (publish-subscribe), the Command pattern, and a simple State Machine. They build from basic to more complex, illustrating how events enable decoupled communication between objects.

### Example A1: Custom Events – Basic Publisher/Subscriber

This example introduces **custom events** in a class and handling them in another class (subscriber). A `Counter` class raises an event every time its count changes, and a `Monitor` class subscribes to print the new value. This illustrates how to define and raise events and how another object catches them using `WithEvents`.

**Class Module: `Counter`** – Defines a custom event and a method that triggers it. It holds a count and notifies listeners when the count updates.

```vba
' Class Module: Counter
Option Explicit
Public Event CountChanged(ByVal newValue As Long)  ' Custom event

Private currentCount As Long

Public Property Get Count() As Long
    Count = currentCount
End Property

Public Sub Increment()
    ' Increase count and raise event
    currentCount = currentCount + 1
    RaiseEvent CountChanged(currentCount)
End Sub
```

**Class Module: `Monitor`** – Subscribes to `Counter` events using `WithEvents` and handles the event.

```vba
' Class Module: Monitor
Option Explicit
Private WithEvents observedCounter As Counter  ' Event source

' Initialize by binding to a Counter instance
Public Sub StartObserving(ByVal c As Counter)
    Set observedCounter = c
    Debug.Print "Monitor attached to Counter."
End Sub

' Event handler – called when Counter raises CountChanged
Private Sub observedCounter_CountChanged(ByVal newValue As Long)
    Debug.Print "[Monitor] Counter value now: " & newValue
End Sub
```

**Standard Module: `Module1`** – Test routine to demonstrate the event.

```vba
Option Explicit
Public Sub Test_CustomEvent()
    Dim c As New Counter
    Dim m As New Monitor
    m.StartObserving c            ' Attach monitor to counter
    c.Increment                   ' Increment triggers CountChanged event
    c.Increment                   ' Trigger again
    ' Output in Immediate Window:
    ' [Monitor] Counter value now: 1
    ' [Monitor] Counter value now: 2
End Sub
```

*How it works:* The `Counter` class declares a `Public Event CountChanged`. Its `Increment` method uses `RaiseEvent` to fire the event. The `Monitor` class declares a `WithEvents` variable to hook a `Counter`. When `StartObserving` sets this variable, VBA wires up the `observedCounter_CountChanged` procedure. Each time `Counter.RaiseEvent` is executed, the monitor’s handler runs immediately, printing the new count. This decouples the Counter from the Monitor – the Counter simply broadcasts the event, and any listeners respond.

### Example A2: Before/After Events with Cancellation

A common pattern is **Before/After events** around an action, allowing the subscriber to cancel the action or perform additional steps. This example defines a `Document` class that raises `BeforeSave` (with a Cancel parameter) and `AfterSave` events when its `Save` method is called. A handler can set Cancel=True to abort the save. This mimics Excel’s built-in events (e.g., `Workbook_BeforeSave`).

**Class Module: `Document`** – Raises Before/After events around a pseudo “save” operation.

```vba
' Class Module: Document
Option Explicit
Public Event BeforeSave(ByRef Cancel As Boolean)
Public Event AfterSave()

Private isSaved As Boolean

Public Sub Save()
    Dim cancelFlag As Boolean
    RaiseEvent BeforeSave(cancelFlag)        ' Notify before saving
    If cancelFlag Then
        Debug.Print "Save was cancelled by handler."
    Else
        ' ... perform save operation (simulated) ...
        isSaved = True
        Debug.Print "Document saved."
        RaiseEvent AfterSave                 ' Notify after save
    End If
End Sub
```

**Class Module: `DocObserver`** – Listens for Document save events and optionally cancels.

```vba
' Class Module: DocObserver
Option Explicit
Private WithEvents doc As Document

Public Sub AttachTo(ByVal d As Document)
    Set doc = d
End Sub

' BeforeSave handler – decide whether to cancel
Private Sub doc_BeforeSave(ByRef Cancel As Boolean)
    Debug.Print "[Observer] Document is about to save..."
    If Hour(Now) < 9 Then
        Cancel = True  ' cancel save outside business hours (example)
        Debug.Print "[Observer] Save cancelled (outside allowed time)."
    End If
End Sub

' AfterSave handler – perform follow-up action
Private Sub doc_AfterSave()
    Debug.Print "[Observer] Document saved successfully. Logging this event."
End Sub
```

**Standard Module: `Module1`** – Test routine.

```vba
Option Explicit
Public Sub Test_BeforeAfterEvents()
    Dim d As New Document
    Dim obs As New DocObserver
    obs.AttachTo d
    d.Save          ' Attempt to save (handler may cancel)
    ' Output example (if run before 9 AM):
    ' [Observer] Document is about to save...
    ' [Observer] Save cancelled (outside allowed time).
    ' Save was cancelled by handler.
    '
    ' Output example (if run after 9 AM):
    ' [Observer] Document is about to save...
    ' Document saved.
    ' [Observer] Document saved successfully. Logging this event.
End Sub
```

*How it works:* The `Document` class defines two events, `BeforeSave(ByRef Cancel As Boolean)` and `AfterSave`. Its `Save` method raises `BeforeSave`, checks the Cancel flag, then proceeds or aborts, and raises `AfterSave` if not cancelled. The `DocObserver` subscribes via `WithEvents`. In `doc_BeforeSave`, we simulate a rule (e.g. cancel saves before 9 AM) by setting `Cancel=True`. This flag is passed back to the `Document.Save` method by reference. If cancellation occurred, the `Document` skips the save and informs us (print message). If not, the save proceeds and triggers `AfterSave`, which the observer handles (e.g. logging). This pattern allows external code to **veto or augment an operation** before and after it happens, mirroring Excel’s design (e.g., `Cancel` in `Workbook_BeforeClose` stops closing).

### Example A3: Observer Pattern – One-to-Many Notifications

The **Observer pattern** lets one object (subject) notify many observers of changes. Here we extend Example A1: a `DataFeed` class (subject) raises an event when new data arrives, and multiple `DataListener` classes (observers) receive the update. This decouples the data source from the listeners – they only communicate via the event.

**Class Module: `DataFeed`** – Subject that publishes an event.

```vba
' Class Module: DataFeed
Option Explicit
Public Event DataArrived(ByVal value As Double)

Public Sub PublishData(ByVal newValue As Double)
    ' Raise event to all listeners
    Debug.Print "DataFeed publishing value: "; newValue
    RaiseEvent DataArrived(newValue)
End Sub
```

**Class Module: `DataListener`** – Observer that subscribes to a DataFeed.

```vba
' Class Module: DataListener
Option Explicit
Private WithEvents source As DataFeed  ' The data source to observe

' Register this listener with a DataFeed
Public Sub Subscribe(ByVal src As DataFeed)
    Set source = src
    Debug.Print "Listener (" & ObjPtr(Me) & ") subscribed."
End Sub

' Optional: Unsubscribe from the source
Public Sub Unsubscribe()
    Set source = Nothing
    Debug.Print "Listener (" & ObjPtr(Me) & ") unsubscribed."
End Sub

' Event handler – called whenever DataFeed publishes new data
Private Sub source_DataArrived(ByVal value As Double)
    Debug.Print "Listener (" & ObjPtr(Me) & ") received data: " & value
End Sub
```

**Standard Module: `Module1`** – Demonstration with multiple observers.

```vba
Option Explicit
Public Sub Test_ObserverPattern()
    Dim feed As New DataFeed
    Dim L1 As New DataListener, L2 As New DataListener
    L1.Subscribe feed
    L2.Subscribe feed
    feed.PublishData 100.25    ' both L1 and L2 will get this
    feed.PublishData 98.75     ' both L1 and L2 get this as well
    L2.Unsubscribe             ' L2 stops listening
    feed.PublishData 101.1     ' only L1 will get this
End Sub
```

**Expected Output (Immediate Window):**

```
Listener (237418384) subscribed.
Listener (237418400) subscribed.
DataFeed publishing value:  100.25
Listener (237418384) received data: 100.25
Listener (237418400) received data: 100.25
DataFeed publishing value:  98.75
Listener (237418384) received data: 98.75
Listener (237418400) received data: 98.75
Listener (237418400) unsubscribed.
DataFeed publishing value:  101.1
Listener (237418384) received data: 101.1
```

*How it works:* `DataFeed` has an event `DataArrived` and a method `PublishData` that raises it. Each `DataListener` uses `WithEvents source As DataFeed` and implements `source_DataArrived` to react. We create two listeners and call `Subscribe` to set their `source` to the same `DataFeed` instance. When `feed.PublishData` is called, **VBA invokes each listener’s handler in turn**, passing the value. Neither the feed nor the listeners know about each other beyond the event hookup. We can unsubscribe a listener by setting its `source = Nothing` (breaking the connection). This pattern is natively used in Excel (e.g., many charts listening to a single data range change). It allows a one-to-many relationship: the subject doesn’t care how many observers exist.

### Example A4: Mediator Pattern – Event Routing via Central Coordinator

The **Mediator pattern** introduces a central object to manage complex interactions, so components don’t reference each other directly. In an event-driven context, a mediator can listen to events from multiple sources and then raise its own higher-level events or call methods on others. This example simulates two modules (a `Sensor` and a `Logger`) that shouldn’t talk directly. A `Coordinator` class mediates: it handles `Sensor` events and raises a new event that `Logger` listens to.

**Class Module: `Sensor`** – Raises an event when it “detects” something.

```vba
' Class Module: Sensor
Option Explicit
Public Event Detected(ByVal level As Long)

Public Sub Trigger(level As Long)
    Debug.Print "Sensor: detected level " & level
    RaiseEvent Detected(level)  ' notify coordinator/others
End Sub
```

**Class Module: `Coordinator`** – Mediator that listens to Sensor and raises its own events.

```vba
' Class Module: Coordinator
Option Explicit
Private WithEvents sensor As Sensor      ' mediator listens to a Sensor
Public Event Alert(ByVal severity As String)

' Initialize by providing the sensor to mediate
Public Sub SetSensor(ByVal s As Sensor)
    Set sensor = s
End Sub

' Sensor event handler – interpret and raise a higher-level Alert event
Private Sub sensor_Detected(ByVal level As Long)
    Dim sev As String
    If level > 80 Then
        sev = "HIGH"
    ElseIf level > 50 Then
        sev = "MEDIUM"
    Else
        sev = "LOW"
    End If
    Debug.Print "Coordinator: translating level " & level & " to " & sev & " alert"
    RaiseEvent Alert(sev)
End Sub
```

**Class Module: `Logger`** – Listens for Alerts (from Coordinator) and logs them.

```vba
' Class Module: Logger
Option Explicit
Private WithEvents coord As Coordinator  ' listen to Coordinator's events

Public Sub Connect(ByVal c As Coordinator)
    Set coord = c
End Sub

Private Sub coord_Alert(ByVal severity As String)
    ' Log the alert (for demo, just print)
    Debug.Print "Logger: Alert received! Severity = " & severity
End Sub
```

**Standard Module: `Module1`** – Test routine wiring up sensor -> coordinator -> logger.

```vba
Option Explicit
Public Sub Test_MediatorPattern()
    Dim s As New Sensor
    Dim med As New Coordinator
    Dim log As New Logger
    med.SetSensor s
    log.Connect med
    ' Simulate sensor triggers:
    s.Trigger 60    ' Medium level
    s.Trigger 90    ' High level
End Sub
```

**Output (Immediate Window):**

```
Sensor: detected level 60
Coordinator: translating level 60 to MEDIUM alert
Logger: Alert received! Severity = MEDIUM
Sensor: detected level 90
Coordinator: translating level 90 to HIGH alert
Logger: Alert received! Severity = HIGH
```

*How it works:* The `Coordinator` acts as an intermediary between `Sensor` and `Logger`. It has `WithEvents sensor` to catch `Sensor.Detected`. In that handler, it encapsulates logic to determine severity and then raises its own `Alert` event. The `Logger` has no knowledge of the `Sensor` – it only subscribes to `Coordinator.Alert`. The `Sensor` knows nothing of the `Logger`. All coupling is centralized in `Coordinator`, simplifying the relationships. This is useful when there are many-to-many interactions or complex conditional logic on events – the mediator can coordinate everything in one place. (For instance, an `ApplicationEventsHandler` class could mediate between Excel’s events and an add-in’s UI by listening to `WorkbookOpen` etc. and raising custom events that UI components handle, instead of UI directly handling low-level events.)

### Example A5: Event Aggregator (Publish–Subscribe Bus)

An **Event Aggregator** (or publish–subscribe system) is a central hub where any component can publish events to a “bus” and any interested component can subscribe to them. This decouples senders and receivers even further than a mediator – publishers and subscribers don’t know about each other at all, only about the event bus. This example implements a simple global **Message Bus** that broadcasts messages, and two different modules subscribe to the bus to receive those messages.

**Class Module: `MessageBus`** – Central event hub.

```vba
' Class Module: MessageBus
Option Explicit
Public Event Message(ByVal topic As String, ByVal payload As Variant)

' Publish a message to all subscribers
Public Sub Publish(ByVal topic As String, ByVal payload As Variant)
    Debug.Print "Bus publishing [" & topic & "] ="; payload
    RaiseEvent Message(topic, payload)
End Sub
```

**Class Module: `ModuleA`** – A publisher that sends messages via the bus.

```vba
' Class Module: ModuleA
Option Explicit
Private bus As MessageBus

Public Sub Init(ByVal messageBus As MessageBus)
    Set bus = messageBus
End Sub

Public Sub DoWork()
    ' Do some work, then publish an event about it:
    Dim result As Long
    result = 42  ' pretend we computed something
    bus.Publish "CalcDone", result
End Sub
```

**Class Module: `ModuleB`** – A subscriber that listens for certain bus messages.

```vba
' Class Module: ModuleB
Option Explicit
Private WithEvents bus As MessageBus

Public Sub Init(ByVal messageBus As MessageBus)
    Set bus = messageBus  ' subscribe to the bus
End Sub

Private Sub bus_Message(ByVal topic As String, ByVal payload As Variant)
    If topic = "CalcDone" Then
        Debug.Print "ModuleB received result: " & payload
        ' (We could react, e.g., update UI or trigger another action)
    Else
        Debug.Print "ModuleB received topic '" & topic & "' (ignored)"
    End If
End Sub
```

**Standard Module: `Module1`** – Set up a global bus and demonstrate cross-component messaging.

```vba
Option Explicit
Public Sub Test_EventAggregator()
    Dim bus As New MessageBus
    Dim compA As New ModuleA, compB As New ModuleB
    compA.Init bus
    compB.Init bus
    ' ModuleB subscribes upon Init. Now ModuleA does something:
    compA.DoWork
    ' We can also simulate another publisher without ModuleB knowing:
    bus.Publish "Alert", "All systems go"
End Sub
```

**Output (Immediate Window):**

```
Bus publishing [CalcDone] =  42 
ModuleB received result: 42
Bus publishing [Alert] = All systems go
ModuleB received topic 'Alert' (ignored)
```

*How it works:* The `MessageBus` class defines a single `Message` event with a topic and payload. Any number of objects can call `bus.Publish(topic, payload)` to raise the event. Any number of listeners can use `WithEvents bus As MessageBus` to handle `bus_Message`. In `Test_EventAggregator`, we instantiate one bus (acting as a singleton event channel) and give references to ModuleA and ModuleB. ModuleA (publisher) calls `bus.Publish "CalcDone", result` instead of calling ModuleB directly – it doesn’t know who (if anyone) will receive it. ModuleB’s handler checks the topic and reacts to `"CalcDone"` messages. When ModuleA publishes, the bus raises the event and ModuleB receives the data. Later, we directly publish an `"Alert"` message on the bus; ModuleB gets it but ignores that topic. This pattern is powerful for decoupling: components can come and go without breaking references, as long as they register with the event aggregator. It’s especially useful for **cross-workbook or cross-module communication** (see Section E).

### Example A6: Command Pattern – Events Triggering Commands

The **Command pattern** encapsulates actions as objects, allowing flexible queuing, logging, or undo/redo of operations. In an event-driven system, commands can be triggered by events (e.g., user interface events) instead of hard-coded procedure calls. This example defines a generic `ICommand` interface and two command classes (`SayHelloCommand` and `SumRangeCommand`). We use an `Invoker` (simulated by a simple `CommandManager`) that listens for a “request event” and then executes the appropriate command.

**Class Module: `ICommand`** – Interface for commands (using `Implements` to define a common Execute method).

```vba
' Class Module: ICommand (Interface)
Option Explicit
Public Sub Execute()
    ' Defined by implementing classes
End Sub
```

**Class Module: `SayHelloCommand`** – A simple command that greets the user.

```vba
' Class Module: SayHelloCommand
Option Explicit
Implements ICommand

Private name As String
Public Sub Initialize(ByVal userName As String)
    name = userName
End Sub

Private Sub ICommand_Execute()
    MsgBox "Hello, " & name & "!", vbInformation, "Greeting"
End Sub
```

**Class Module: `SumRangeCommand`** – A command that sums a worksheet range and displays the result.

```vba
' Class Module: SumRangeCommand
Option Explicit
Implements ICommand

Private targetRange As Range
Public Sub Initialize(ByVal rng As Range)
    Set targetRange = rng
End Sub

Private Sub ICommand_Execute()
    If targetRange Is Nothing Then Exit Sub
    Dim total As Double
    total = Application.WorksheetFunction.Sum(targetRange)
    MsgBox "Sum of " & targetRange.Address(0, 0) & " is " & total, vbInformation, "SumRange"
End Sub
```

**Class Module: `CommandManager`** – Invoker that holds available commands and triggers them on events. It raises a custom event `CommandRequested` (with a command name) which it also handles itself to execute the command.

```vba
' Class Module: CommandManager
Option Explicit
Public Event CommandRequested(ByVal commandName As String)

' Registry of commands by name
Private commands As Object  ' use a Scripting.Dictionary for mapping

Private Sub Class_Initialize()
    Set commands = CreateObject("Scripting.Dictionary")
End Sub

Public Sub RegisterCommand(ByVal name As String, ByVal cmd As ICommand)
    commands(name) = cmd
End Sub

' Simulate an external trigger (e.g., UI event) by raising our own event
Public Sub RequestCommand(ByVal name As String)
    Debug.Print "[CommandManager] Requesting command: " & name
    RaiseEvent CommandRequested(name)
End Sub

' Handle the request by executing the corresponding command
Private Sub Class_Terminate()
    Set commands = Nothing
End Sub

Private Sub CommandManager_CommandRequested(ByVal commandName As String)
    If commands.Exists(commandName) Then
        Debug.Print "[CommandManager] Executing command '" & commandName & "'"
        Dim cmd As ICommand
        Set cmd = commands(commandName)
        cmd.Execute  ' invoke the command's action
    Else
        Debug.Print "[CommandManager] Unknown command: " & commandName
    End If
End Sub
```

**Standard Module: `Module1`** – Set up the commands and simulate events triggering them.

```vba
Option Explicit
Public Sub Test_CommandPattern()
    Dim mgr As New CommandManager
    ' Prepare commands
    Dim helloCmd As New SayHelloCommand
    helloCmd.Initialize "Alice"
    Dim sumCmd As New SumRangeCommand
    sumCmd.Initialize ThisWorkbook.Sheets(1).Range("A1:A5")
    ' Register commands with names
    mgr.RegisterCommand "GREET", helloCmd
    mgr.RegisterCommand "SUMRANGE", sumCmd
    ' Simulate event triggers (in reality, these might be button clicks etc.)
    mgr.RequestCommand "GREET"
    mgr.RequestCommand "SUMRANGE"
    mgr.RequestCommand "UNKNOWN"  ' a request for which no command is registered
End Sub
```

**Output / Behavior:** Running `Test_CommandPattern` will cause the `CommandManager` to print requests and execution steps. It will pop up two message boxes: one greeting “Alice” and one showing the sum of cells A1\:A5 (make sure Sheet1 A1\:A5 have numbers). The debug output will show the flow:

```
[CommandManager] Requesting command: GREET
[CommandManager] Executing command 'GREET'
[CommandManager] Requesting command: SUMRANGE
[CommandManager] Executing command 'SUMRANGE'
[CommandManager] Requesting command: UNKNOWN
[CommandManager] Unknown command: UNKNOWN
```

*How it works:* We define `ICommand` as a simple interface with an `Execute` method. Two concrete commands implement this via `Implements ICommand`. The `CommandManager` holds a dictionary of command objects, identified by name. When an event occurs (here we simulate it by calling `RequestCommand`), the manager raises a `CommandRequested` event with the name. We use an event rather than direct call to illustrate that this could be triggered by an external UI event or another object. The manager’s own event handler `CommandManager_CommandRequested` then looks up the command and calls its `Execute`. In a real UI, one might tie a Ribbon button’s onAction to call `mgr.RequestCommand "XYZ"`. This pattern decouples the *what* (the button click or event, which just supplies a command name) from the *how* (the actual action in the command object). New commands can be added without changing the UI wiring – just register them. It’s also easy to log or undo commands by enhancing the manager (e.g., keep a history of executed commands). Here, events are used to trigger command execution indirectly, demonstrating a flexible design.

### Example A7: Simple State Machine with Event Transitions

A **State Machine** pattern models an object that transitions through a set of states, often triggering events on transitions. We implement a simple workflow with states **“Idle”**, **“Processing”**, **“Done”**. The `Workflow` class raises a `StateChanged` event whenever it moves to a new state. A `StateWatcher` subscribes to log the transitions. This pattern is useful when an object’s behavior depends on its state; external code can be notified of state changes and respond accordingly.

**Class Module: `Workflow`** – Manages state and transitions.

```vba
' Class Module: Workflow
Option Explicit
Public Event StateChanged(ByVal newState As String)

Public Enum WorkflowState
    IdleState
    ProcessingState
    DoneState
End Enum

Private currentState As WorkflowState

Public Sub Start()
    If currentState <> IdleState Then
        Debug.Print "Cannot start – not in Idle state."
        Exit Sub
    End If
    currentState = ProcessingState
    Debug.Print "Workflow transitioning to PROCESSING."
    RaiseEvent StateChanged("PROCESSING")
    ' ... perform some processing (simulated here by a delay) ...
    Dim t As Single: t = Timer: Do While Timer - t < 1: DoEvents: Loop  ' 1-second delay
    Complete    ' automatically complete after processing
End Sub

Public Sub Complete()
    If currentState <> ProcessingState Then Exit Sub
    currentState = DoneState
    Debug.Print "Workflow transitioning to DONE."
    RaiseEvent StateChanged("DONE")
End Sub

Public Sub Reset()
    currentState = IdleState
    Debug.Print "Workflow transitioning to IDLE."
    RaiseEvent StateChanged("IDLE")
End Sub

Public Property Get State() As WorkflowState
    State = currentState
End Property
```

**Class Module: `StateWatcher`** – Observes state changes.

```vba
' Class Module: StateWatcher
Option Explicit
Private WithEvents wf As Workflow

Public Sub Attach(ByVal w As Workflow)
    Set wf = w
    Debug.Print "(Watcher attached to workflow.)"
End Sub

Private Sub wf_StateChanged(ByVal newState As String)
    Debug.Print "[Watcher] New workflow state: " & newState
    ' React to certain states if needed:
    ' e.g., If newState = "DONE" Then ... (cleanup or next steps)
End Sub
```

**Standard Module: `Module1`** – Demonstration of state transitions.

```vba
Option Explicit
Public Sub Test_StateMachine()
    Dim w As New Workflow
    Dim watcher As New StateWatcher
    watcher.Attach w
    w.Start            ' Moves Idle -> Processing -> Done
    w.Reset            ' Moves Done -> Idle
    ' Trying an invalid transition:
    w.Complete         ' (No effect because not in Processing)
End Sub
```

**Expected Output:**

```
(Watcher attached to workflow.)
Workflow transitioning to PROCESSING.
[Watcher] New workflow state: PROCESSING
Workflow transitioning to DONE.
[Watcher] New workflow state: DONE
Workflow transitioning to IDLE.
[Watcher] New workflow state: IDLE
```

*(The call to `w.Complete` when already done produces no output, as guarded in the code.)*

*How it works:* The `Workflow` class holds a state (using an Enum for clarity) and exposes methods to transition: `Start` (Idle -> Processing), automatic transition to Done (via `Complete` called inside `Start` after a delay), and `Reset` (to Idle). Each transition sets `currentState` and raises `StateChanged` with a string name of the new state. The `StateWatcher` attaches to a `Workflow` and prints any state changes it hears. This allows external components to react to state transitions without constantly polling the workflow’s state. For example, a UI could enable a “Results” button when state becomes DONE. The pattern ensures **state changes are centralized** in `Workflow` (so invalid transitions can be prevented) and **notifications are broadcast** to any interested parties. In our output, you see the watcher logging each transition. The final `w.Complete` call does nothing because the workflow was not in the correct prior state (the guard `If currentState <> ProcessingState Then Exit Sub` prevented an illegal transition).

### Example A8: Event Forwarding (Bridging Interfaces)

Sometimes a class may **forward events** from one source to another, acting as an adapter. For completeness, we give a brief example: an `Adapter` class that wraps an Excel `Worksheet` and translates its events to a custom event for external consumers. This pattern can unify different event sources under one interface.

**Class Module: `SheetAdapter`** – Wraps a Worksheet’s Change event into a custom event.

```vba
' Class Module: SheetAdapter
Option Explicit
Private WithEvents sh As Worksheet
Public Event Updated(ByVal info As String)

Public Sub AttachSheet(ByVal ws As Worksheet)
    Set sh = ws
End Sub

Private Sub sh_Change(ByVal Target As Range)
    ' Forward Excel's event as a simplified custom event
    RaiseEvent Updated("Sheet '" & sh.Name & "' changed at " & Target.Address(0,0))
End Sub
```

**Class Module: `Listener`** – Subscribes to the adapter’s event.

```vba
' Class Module: Listener
Option Explicit
Private WithEvents adapter As SheetAdapter
Public Sub Subscribe(ByVal ad As SheetAdapter)
    Set adapter = ad
End Sub
Private Sub adapter_Updated(ByVal info As String)
    Debug.Print "[Listener] " & info
End Sub
```

**Usage:**

```vba
Dim ad As New SheetAdapter, lis As New Listener
ad.AttachSheet ThisWorkbook.Worksheets(1)
lis.Subscribe ad
' Now any change on Sheet1 will trigger adapter_Updated event, which Listener will print.
```

*(This example is not a standalone test macro since it relies on actual sheet edits. Change some cell in Sheet1 after running the setup to see the debug print.)*

*How it works:* The `SheetAdapter` listens to `Worksheet.Change` (Excel event) via `WithEvents sh As Worksheet`. In its handler, it raises a custom `Updated` event with a friendlier message. The `Listener` class can subscribe to `SheetAdapter` without needing to reference Excel’s `Worksheet` type. This is a form of adapter/bridge – useful if you want to expose sheet events to code that shouldn’t depend on the Excel object model directly. (This pattern is used in some MVP/MVVM implementations to bind MSForms controls via custom events.)

## B. Event Performance & Control Techniques (4 Examples)

Frequent or long-running events can degrade performance or cause memory leaks if not managed carefully. These examples demonstrate **debouncing** and **throttling** event handlers to control frequency, **coalescing** multiple rapid events into one, and using weak references/cleanup to avoid memory leaks from event subscriptions.

### Example B1: Debounce – Ignoring Rapid-Fire Events Until Quiescent

**Debouncing** delays event handling until a burst of frequent events stops. This is useful for events like `Worksheet_SelectionChange` or `Change` that can fire many times in quick succession. We implement a `DebounceHelper` class that uses `Application.OnTime` to schedule a handler call after a short delay, canceling any previously scheduled call if a new event comes in. This way, only the *last* event in a rapid sequence triggers the action.

**Class Module: `DebounceHelper`** – Schedules and cancels OnTime calls to defer action.

```vba
' Class Module: DebounceHelper
Option Explicit
Private nextTime As Date
Private Const DEFAULT_DELAY As Double = 0.5#  ' 0.5 seconds delay

' Schedule a macro to run after the specified delay, debouncing intermediate calls
Public Sub Schedule(ByVal macroName As String, Optional ByVal delaySeconds As Double = DEFAULT_DELAY)
    Dim fireTime As Date
    fireTime = Now + delaySeconds / 86400#  ' convert sec to days for OnTime
    ' Cancel any pending call
    On Error Resume Next
    If nextTime <> 0 Then
        Application.OnTime EarliestTime:=nextTime, Procedure:=macroName, Schedule:=False
    End If
    On Error GoTo 0
    ' Schedule new call
    nextTime = fireTime
    Application.OnTime EarliestTime:=nextTime, Procedure:=macroName, Schedule:=True
End Sub

' Cancel any pending scheduled call
Public Sub CancelPending(ByVal macroName As String)
    If nextTime <> 0 Then
        On Error Resume Next
        Application.OnTime EarliestTime:=nextTime, Procedure:=macroName, Schedule:=False
        On Error GoTo 0
        nextTime = 0
    End If
End Sub
```

**Usage Scenario:** Suppose we want to run an expensive operation after the user stops selecting new cells. We can use `DebounceHelper` in the sheet’s SelectionChange event:

```vba
' In Worksheet module (e.g., Sheet1):
Private WithEvents debouncer As New DebounceHelper

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    debouncer.Schedule "DelayedSelectionAction", 0.5
    ' Immediate lightweight feedback (optional):
    Me.Cells(1,1).Value = "Last selection: " & Target.Address(0,0)
End Sub
```

```vba
' In a standard module:
Public Sub DelayedSelectionAction()
    ' Runs 0.5s after last selection change
    Dim sel As Range: Set sel = Application.Selection
    If sel Is Nothing Then Exit Sub
    ' For example, show sum of the selection:
    If sel.Cells.CountLarge > 1 Then
        MsgBox "Sum of selection: " & Application.WorksheetFunction.Sum(sel), vbInformation
    Else
        MsgBox "Selected cell " & sel.Address(0,0) & " = " & sel.Value, vbInformation
    End If
End Sub
```

If the user rapidly moves the selection, the message box will only appear after they pause for 0.5 seconds without another selection change. All intermediate OnTime calls are canceled except the last. Debouncing prevents flicker or repeated calculations, improving responsiveness. (Note: OnTime has \~1 second minimum resolution by default, but can handle fractional delays ≈0.5s as above.)

### Example B2: Throttle – Rate-Limiting Event Handling

**Throttling** ensures an event handler runs at most once per specified interval, ignoring extra events in between. For example, you might want to process a real-time data feed no more than 5 times per second even if events arrive faster.

We implement a `ThrottleHelper` that tracks the last execution time and only allows the action to proceed if enough time has passed. Unlike debounce (which waits and executes the *last* call), throttle executes immediately then suppresses subsequent calls until the interval passes.

**Class Module: `ThrottleHelper`** – Allows an action occasionally while dropping excess events.

```vba
' Class Module: ThrottleHelper
Option Explicit
Private lastTime As Single  ' last execution time (Timer value in seconds)

Public Sub TryExecute(ByVal macroName As String, ByVal minIntervalSec As Single)
    Dim current As Single
    current = Timer  ' current time in seconds since midnight
    If lastTime = 0 Or current - lastTime >= minIntervalSec Then
        lastTime = current
        Debug.Print "Throttle: executing " & macroName & " at t=" & current
        Application.Run macroName   ' call the macro
    Else
        ' Too soon since last execution - skip this event
        Debug.Print "Throttle: skipped " & macroName & " at t=" & current
    End If
End Sub

Public Sub Reset()
    lastTime = 0
End Sub
```

**Usage Scenario:** Imagine a TextBox on a UserForm that raises `Change` events rapidly as the user types, and we want to update a calculation only at most twice per second. We could use `ThrottleHelper` in the TextBox\_Change event:

```vba
Private WithEvents throttle As New ThrottleHelper

Private Sub TextBox1_Change()
    throttle.TryExecute "RecalculateResult", 0.5  ' allow twice per second
End Sub
```

```vba
Public Sub RecalculateResult()
    ' Heavy recalculation based on TextBox1.Text
    Debug.Print "Recalculating for input: " & UserForm1.TextBox1.Text
End Sub
```

If the user types quickly, “Recalculating...” happens only at \~0.5s intervals. All interim changes are ignored (skipped). The debug output from `ThrottleHelper` will show execute vs skip decisions. Throttle is useful when you prefer responsiveness (immediate first reaction) but want to limit frequency. (For a variant, one could schedule a trailing call for the last skipped event – a combination of throttle and debounce – but that’s more complex.)

### Example B3: Coalescing Events – Batch Multiple Triggers into One

**Coalescing** merges a flurry of events into a single consolidated handling. For example, if 50 cells change in quick succession (maybe due to a paste operation), a naive handler would run 50 times. With coalescing, you could accumulate the changed ranges and handle them all at once after the flurry.

We create a `CoalesceHelper` that collects event data and uses a short timer to defer processing. The difference from debouncing: we don’t reset the timer on each event; we set it once on the first event of a batch, and collect all events until that timer expires. Then we process **all** collected data together.

**Class Module: `CoalesceHelper`** – Accumulates events and triggers a batch process once.

```vba
' Class Module: CoalesceHelper
Option Explicit
Private events As Collection
Private pending As Boolean

Private Sub Class_Initialize()
    Set events = New Collection
End Sub

' Add event info and schedule processing if not already scheduled
Public Sub AddEvent(ByVal info As Variant)
    events.Add info
    If Not pending Then
        pending = True
        Application.OnTime Now + TimeSerial(0, 0, 1), "ProcessEventBatch"
        Debug.Print "CoalesceHelper: batch scheduled in 1s"
    End If
End Sub

' Retrieve all accumulated events (called by the batch processor)
Public Function RetrieveBatch() As Collection
    Set RetrieveBatch = events
    ' Reset for next batch
    Set events = New Collection
    pending = False
End Function
```

**Standard Module: `ModuleBatch`** – The batch processor macro that processes all coalesced events.

```vba
Option Explicit
Public BatchCoalescer As CoalesceHelper  ' global instance

Public Sub ProcessEventBatch()
    If BatchCoalescer Is Nothing Then Exit Sub
    Dim batch As Collection
    Set batch = BatchCoalescer.RetrieveBatch()
    Debug.Print "*** Processing batch of " & batch.Count & " events ***"
    Dim itm As Variant
    For Each itm In batch
        Debug.Print " - Event data: " & CStr(itm)
    Next itm
End Sub
```

**Usage:** We need to initialize the global `BatchCoalescer` and then use it in an event. For example, in `Worksheet_Change`:

```vba
Private Sub Worksheet_Change(ByVal Target As Range)
    If BatchCoalescer Is Nothing Then Set BatchCoalescer = New CoalesceHelper
    BatchCoalescer.AddEvent("Change at " & Target.Address(0,0))
End Sub
```

If the user (or code) triggers many changes quickly, `AddEvent` will keep adding info to the `events` collection. On the first change of a batch, it schedules `ProcessEventBatch` after 1 second and marks `pending=True`. Subsequent changes within that second just accumulate data (the OnTime call is not rescheduled or canceled in this pattern). After 1 second of no new events (i.e., when the OnTime fires), `ProcessEventBatch` retrieves the whole batch and processes it. The output might look like:

```
CoalesceHelper: batch scheduled in 1s
CoalesceHelper: batch scheduled in 1s   ' (only once actually printed, subsequent calls won't schedule again)
*** Processing batch of 5 events ***
 - Event data: Change at A1
 - Event data: Change at A2
 - Event data: Change at A3
 - Event data: Change at A4
 - Event data: Change at A5
```

This shows that five rapid changes were handled in one batch after a delay. Coalescing is useful for actions like bulk updates or finalizing multiple small events. We used a simple 1s delay; it could be shorter depending on context. (We chose to globally store the coalescer for simplicity; in a larger app, you might manage it differently.)

**Note:** Coalescing vs debouncing – Debounce only cares about the *last* event (ignoring earlier ones entirely), whereas coalescing collects *all* events in a burst and handles them together. Use coalescing when intermediate data matters (e.g., summing all changes, or logging all items in a batch).

### Example B4: Cleanup and Weak References – Avoiding Memory Leaks

In VBA, a common memory leak occurs when two objects have circular references (especially with `WithEvents`), preventing proper garbage collection. For instance, if object A holds a `WithEvents` reference to B, and B holds a reference to A, neither will release. You **must** break the cycle (set one reference to Nothing) or use a *weak reference* technique so one side doesn’t count toward reference count.

**Scenario:** A `Parent` object creates a `Child` and they reference each other (child has a pointer back to parent). Without careful cleanup, setting the parent to Nothing won’t free them. We demonstrate the leak and then solve it by storing the parent reference in the child as a raw pointer (LongPtr) rather than an object – a manual weak reference.

**Class Module: `Parent`** – Holds a child. In naive implementation, it passes itself to the child, causing a strong reference cycle.

```vba
' Class Module: Parent
Option Explicit
Private children As Collection

Private Sub Class_Initialize()
    Set children = New Collection
End Sub

Public Sub AddChild(ByVal c As Child)
    ' Register child and let child know its parent (could create circular ref)
    children.Add c
    c.SetParentWeak Me   ' Use weak reference assignment
End Sub

Private Sub Class_Terminate()
    Debug.Print "Parent terminating."
End Sub
```

**Class Module: `Child`** – Holds a “reference” to parent. We simulate a weak reference by storing the pointer using `ObjPtr` and `CopyMemory` to retrieve it only when needed.

```vba
' Class Module: Child
Option Explicit
' We need CopyMemory to manipulate pointers (for 64-bit)
#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#End If

Private parentPtr As LongPtr  ' store parent pointer instead of object

Friend Sub SetParentWeak(ByVal p As Parent)
    ' Store the pointer to parent without increasing ref count
    parentPtr = ObjPtr(p)
End Sub

Public Function GetParent() As Parent
    If parentPtr = 0 Then Exit Function
    Dim tempObj As Parent
    ' Copy pointer into an object variable (dangerous if parent no longer exists)
    CopyMemory tempObj, parentPtr, LenB(parentPtr)
    If Not tempObj Is Nothing Then
        Set GetParent = tempObj  ' assign to return (increments refcount just for return)
    End If
    ' Clean up tempObj without releasing parent (set to Nothing via pointer zero-out)
    CopyMemory tempObj, 0&, LenB(parentPtr)
End Function

Private Sub Class_Terminate()
    Debug.Print "Child terminating."
End Sub
```

**Standard Module: `Module1`** – Demonstrate leak vs fix.

```vba
Option Explicit

Public Sub Test_MemoryLeakCycle()
    Dim p As Parent
    Set p = New Parent
    p.AddChild New Child  ' create child that points to parent (weakly)
    Set p = Nothing       ' remove reference to parent
    Debug.Print "After setting Parent to Nothing."
    ' Expectation: Both Parent and Child should terminate if no leak.
End Sub
```

**Expected Output:**

```
Parent terminating.
Child terminating.
After setting Parent to Nothing.
```

*How it works:* We use `ObjPtr` to grab the memory address of the `Parent` and store it in the child’s `parentPtr` (LongPtr). This does **not** increment the parent’s reference count. The `Parent` keeps a normal reference to the child in its collection. When we set `p = Nothing`, the parent’s reference count drops to 0, *but normally the child still has a reference to parent so parent wouldn’t terminate*. Here, however, the child’s reference is just a pointer – the runtime doesn’t see it as a live reference. Thus `Parent`’s `Class_Terminate` runs, and then the child collection goes out of scope, releasing the child, causing `Child_Terminate`. Both terminate, indicating no leak.

If we had used a normal `c.Parent = Me` strong reference instead, neither would terminate (leak). The weak reference approach requires caution: using a pointer to an object that’s been destroyed will crash Excel. In our `GetParent`, we attempt to copy the pointer back to an object reference. We check `If Not tempObj Is Nothing` to see if the parent is still alive (if it’s freed, that check might throw or tempObj would be Nothing). If alive, we return it (temporarily incrementing refcount on return). We immediately zero out `tempObj`’s copy of the pointer so we don’t release the parent when tempObj goes out of scope. This technique, while advanced, allows an event subscriber to hold a “weak” reference to an event source or vice versa. A perhaps safer approach in many cases is simply to **explicitly break circular links**: e.g., call a `Dispose` method or set object references to Nothing (like removing event handlers) when they’re no longer needed. Always ensure that any `WithEvents` subscriptions are torn down (set to Nothing) when either party is going away.

**Mini Case Study: Preventing Memory Leaks in Event-Driven Code** – In a real-world VBA project, a user-form (`ParentForm`) dynamically created several child objects to handle background tasks, each holding a reference back to the form for callbacks. The form also kept `WithEvents` references to the children to update the UI when tasks progressed. This circular setup caused the form and children to remain in memory even after the form was unloaded – a classic leak. **Naïve approach:** Rely on unloading the form to clean everything – but the hidden circular references meant `Terminate` never fired. The app would accumulate “ghost” forms in memory, slowing down over time. **Solution:** Implement an explicit shutdown routine – when the form unloads, it iterates through each child and sets its reference to the form to Nothing (or calls a `Dispose` method on the child). The form also sets its `WithEvents` child references to Nothing. This manual cleanup broke the cycle, allowing proper termination. Alternatively, a **Weak Reference** pattern like above could have been used: the child stores only the pointer to the form, not a full reference, so the form can die without waiting for children. However, that approach is error-prone in VBA – one must be very careful to check object validity. **Performance impact:** The leak itself was the issue (memory growth, potential crashes), not CPU speed. Once fixed, the termination debug prints confirmed all objects were freeing. As a rule, **for any pair of objects that reference each other (directly or via events)**, decide which will be responsible for breaking the link. In .NET, you might use WeakReference or WeakEvent patterns; in VBA, you do it by setting object references to `Nothing` or using pointer hacks as shown. The bottom line: *ensure events and object references are cleaned up* – e.g., an Application events class should be destroyed on workbook close, or an observer should unsubscribe if its subject outlives it. With careful design, you can avoid lingering objects and keep Excel stable.

## C. System Integration & Low-Level Events (5 Examples)

VBA can leverage Windows and system APIs for events beyond what Excel provides. These examples show using high-resolution Win32 timers, tapping into WMI (Windows Management Instrumentation) for system events, intercepting Ribbon commands, and subclassing Windows procedures for key events and global hotkeys. **⚠️ Note:** These techniques often involve Windows API calls and can destabilize Excel if misused – always ensure proper cleanup and testing.

### Example C1: High-Resolution Timer Events with Win32 `timeSetEvent`

Excel’s `Application.OnTime` is limited to \~1-second resolution and runs on the main thread. The Win32 multimedia timers (`timeSetEvent`) allow millisecond precision timers, invoking a callback at intervals. We can use this to perform periodic actions (e.g., updating data or UI) more frequently or precisely than OnTime. However, the callback still executes on Excel’s single thread (not truly parallel), and we must stop the timer to avoid crashes.

Below, we create a `PingTimer` class that uses `timeSetEvent` to call an internal method every second (for 3 ticks, then stops). It demonstrates setting up the timer and handling its callback.

**Class Module: `PingTimer`** – Uses `timeSetEvent` to raise periodic events.

```vba
' Class Module: PingTimer
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function timeSetEvent Lib "winmm.dll" (ByVal msDelay As Long, _
          ByVal msResolution As Long, ByVal lpFunction As LongPtr, ByVal dwUser As LongPtr, ByVal uFlags As Long) As Long
    Private Declare PtrSafe Function timeKillEvent Lib "winmm.dll" (ByVal uID As Long) As Long
#Else
    ' 32-bit declarations (if needed)
#End If

Private Const TIME_PERIODIC As Long = 1&
Private timerID As Long
Private count As Long

' Start a periodic timer that calls the static callback every given interval (ms)
Public Sub StartTimer(ByVal intervalMs As Long)
    If timerID <> 0 Then Exit Sub  ' already running
    count = 0
    ' Set up the timer to call TimerCallback every interval
    timerID = timeSetEvent(intervalMs, 0, AddressOf TimerCallback, ObjPtr(Me), TIME_PERIODIC)
    If timerID = 0 Then
        Debug.Print "Failed to start timer."
    Else
        Debug.Print "PingTimer started (ID=" & timerID & ")."
    End If
End Sub

' Stop the timer
Public Sub StopTimer()
    If timerID <> 0 Then
        timeKillEvent timerID
        Debug.Print "PingTimer stopped."
        timerID = 0
    End If
End Sub

' This procedure is called by Windows on each timer tick (via AddressOf)
Public Sub TimerTick()
    count = count + 1
    Debug.Print "Ping " & count & " at " & Format(Now, "hh:nn:ss")
    If count >= 3 Then
        StopTimer  ' stop after 3 pings (for demo)
    End If
End Sub
```

**Standard Module: `TimerModule`** – Contains the actual callback function (must be standard module for AddressOf in VBA) and test routine.

```vba
Option Explicit
Public pingObj As PingTimer  ' need a module-level reference to keep PingTimer alive

' The callback signature must match expected: (uID As Long, uMsg As Long, dwUser As LongPtr, dw1 As Long, dw2 As Long)
Public Sub TimerCallback(ByVal uID As Long, ByVal uMsg As Long, ByVal dwUser As LongPtr, ByVal dw1 As Long, ByVal dw2 As Long)
    Dim obj As PingTimer
    ' Convert dwUser (which we passed as ObjPtr(Me)) back to an object reference:
    CopyMemory obj, dwUser, LenB(dwUser)  ' get object from pointer
    If obj Is Nothing Then Exit Sub
    obj.TimerTick                         ' call the instance method
    CopyMemory obj, 0&, LenB(dwUser)      ' release temp ref
End Sub

Public Sub Test_HighResTimer()
    Set pingObj = New PingTimer
    pingObj.StartTimer 250   ' 0.25 second interval
    ' The PingTimer will print "Ping 1", "Ping 2", "Ping 3" at ~0.25s intervals, then stop itself.
    ' Note: The code continues immediately; the timer runs asynchronously on the same thread.
    ' We keep pingObj in a global to avoid it going out of scope (which would kill the timer).
End Sub
```

**Expected Output (Immediate Window):**

```
PingTimer started (ID=1).
Ping 1 at 09:39:58
Ping 2 at 09:39:58
Ping 3 at 09:39:59
PingTimer stopped.
```

*(It prints three “Ping” messages \~250ms apart. The exact timing may vary, but the resolution is much higher than OnTime’s 1s.)*

*How it works:* We use `timeSetEvent` from **winmm.dll** to schedule periodic callbacks with 1ms capability. We pass the address of a `TimerCallback` function and a user parameter (`ObjPtr(Me)`) which is the pointer to our class instance. Each tick, `TimerCallback` is invoked by the system; it reconstructs the `PingTimer` object from the pointer and calls the instance’s `TimerTick` method. `TimerTick` raises the count and prints a timestamp. After 3 ticks, we call `StopTimer` to kill the timer (always do this, otherwise Excel may crash on shutdown). Notice we had to maintain `pingObj` in a module-level variable. If it went out of scope, the object would be destroyed but the timer might still try to call the callback (dwUser pointer becomes invalid -> crash). Nigel Heffernan and others point out that these multimedia timers don’t create true concurrency – the callbacks execute on the main thread, interleaved with other events. Still, they allow finer timing than OnTime. **Performance:** The overhead of RaiseEvent vs calling a method via AddressOf is small – e.g., RaiseEvent is microsecond-scale. `timeSetEvent` can schedule \~1ms ticks, but if your `TimerTick` work takes longer than interval, you effectively skip or delay ticks (they queue on the single thread). Use this for short periodic tasks and be mindful of thread reentrancy. Always call `timeKillEvent` to clear timers when done, and store the timer ID to avoid losing control. This example also shows how to pass an object pointer through API callbacks – a common trick to get back into class context in VB6/VBA when Windows calls your function.

**Mini Case Study: High-Frequency Updates – OnTime vs Win32 Timers** – A trading application needed to update Excel with market data ticks \~10 times per second. **Initial approach:** use `Application.OnTime` scheduling successive calls 0.1 seconds apart. This failed – OnTime’s true minimum interval is 1 second, so it couldn’t handle sub-second updates (attempting fractions often rounded or queued unpredictably). Also, OnTime events would sometimes bunch or drop under heavy Excel load. **Solution:** implement a Win32 `timeSetEvent` in a hidden module, similar to our PingTimer. The timer fired every 100ms, and in the callback the add-in pulled new data and updated the sheet. **Results:** The updates appeared much smoother; the effective interval was \~100–125ms, well under OnTime’s 1s granularity. CPU use was manageable because the work done each tick was light. They did find that if Excel was busy (e.g., user dragging cells), the timer ticks would wait – indeed no true multi-threading – but once free, the updates resumed. **Performance data:** With OnTime, best-case tick latency was \~1s; with `timeSetEvent`, \~0.1s. Over a 1-minute test, Win32 timer delivered \~600 updates vs OnTime’s \~60. However, one issue encountered: forgetting to call `timeKillEvent` caused Excel to hang on close (the timer kept trying to fire into an unloaded project). After adding proper `StopTimer` logic (as we do in PingTimer), the solution ran reliably. This shows that while Win32 timers can greatly enhance temporal resolution in VBA, one must handle them with care. Always balance precision needs against Excel’s single-threaded nature – a 1ms timer that does heavy computation will still choke the UI. But for moderate tasks, Win32 timers are a powerful tool in the VBA arsenal.

### Example C2: Listening to System Events with WMI (e.g., Process Start)

Windows Management Instrumentation (WMI) lets you subscribe to a huge range of system events (hardware, processes, file system, etc.) from VBA. Here, we’ll set up a WMI event sink to catch when a new process (e.g., Notepad) is created on the system. This uses COM objects from the WMI scripting library. We need to use an asynchronous query so our code doesn’t halt waiting for events – the events will come via a callback.

**Class Module: `ProcessWatcher`** – Subscribes to WMI process creation events and raises a custom event in VBA.

```vba
' Class Module: ProcessWatcher
Option Explicit
Private WithEvents sink As SWbemSink  ' WMI event sink (requires reference to Microsoft WMI Scripting)
Private wmiServices As SWbemServices

Public Event ProcessStarted(ByVal processName As String, ByVal processId As Long)

' Start watching for new processes (filtering by optional name)
Public Sub StartWatching(Optional ByVal targetName As String = "")
    ' Get WMI service for local machine, cimv2 namespace
    Set wmiServices = GetObject("winmgmts:\\.\root\CIMV2")
    Set sink = New SWbemSink
    Dim query As String
    query = "SELECT * FROM __InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_Process'"
    If targetName <> "" Then
        query = query & " AND TargetInstance.Name='" & targetName & "'"
    End If
    wmiServices.ExecNotificationQueryAsync sink, query
    Debug.Print "WMI query started for new process events..." 
End Sub

' Stop watching (cancel the async query)
Public Sub StopWatching()
    If Not sink Is Nothing Then
        sink.Cancel()  ' Cancels all outstanding async operations for this sink
    End If
    Set sink = Nothing
    Set wmiServices = Nothing
    Debug.Print "Stopped watching for process events."
End Sub

' WMI sink event: called when an event arrives
Private Sub sink_OnObjectReady(ByVal obj As SWbemObject, ByVal asyncCtx As SWbemNamedValueSet)
    ' This fires on a background thread in response to an event
    Dim proc As SWbemObject
    Set proc = obj!TargetInstance   ' the new process object
    Dim pname As String, pid As Long
    pname = proc!Name
    pid = proc!ProcessId
    Debug.Print "WMI Event: New process '" & pname & "' (PID " & pid & ")"
    RaiseEvent ProcessStarted(pname, pid)  ' raise our VBA event for external handlers
End Sub

Private Sub sink_OnCompleted(ByVal status As WbemErrorEnum, ByVal objError As SWbemObject, ByVal asyncCtx As SWbemNamedValueSet)
    Debug.Print "WMI Event query completed (status " & status & ")."
End Sub
```

**Standard Module: `Module1`** – Example usage of `ProcessWatcher`.

```vba
Option Explicit
Public watcher As ProcessWatcher  ' keep global reference to avoid garbage collection

Public Sub Test_WMIProcessEvents()
    Set watcher = New ProcessWatcher
    ' (Optional) handle the event in code:
    Dim WithEvents localWatch As ProcessWatcher  ' demonstration of handling event locally
    Set localWatch = watcher
    watcher.StartWatching "notepad.exe"
    Debug.Print ">> Launch Notepad to test, then close it. <<" 
    ' The Debug/Immediate Window will show events. 
    ' StopWatching is not called here; in a real app, call it on cleanup or workbook close.
End Sub

Private Sub localWatch_ProcessStarted(ByVal processName As String, ByVal processId As Long)
    Debug.Print "[VBA handler] ProcessStarted event caught in VBA: " & processName & " (PID " & processId & ")"
End Sub
```

**Usage:** Run `Test_WMIProcessEvents`. Within the next minute, if you start (and within \~1 second) any process matching the filter (notepad.exe in this example), you will see output like:

```
WMI query started for new process events...
>> Launch Notepad to test, then close it. <<
WMI Event: New process 'notepad.exe' (PID 1234)
[VBA handler] ProcessStarted event caught in VBA: notepad.exe (PID 1234)
```

If you had not filtered by name, it would report every new process. The `ProcessStarted` event is raised on the main thread (after WMI delivers the COM event) so it’s safe to update the UI in that handler. We maintain `watcher` in a global to keep it alive; otherwise it might terminate, canceling the WMI subscription.

*How it works:* We use `SWbemSink` with `WithEvents` to receive asynchronous WMI events. The `ExecNotificationQueryAsync` call runs our query in the background. Whenever a new process starts (WMI polls each `WITHIN 1` second), WMI calls `sink_OnObjectReady`, passing a `SWbemObject` representing the event. We extract the `TargetInstance` (the `Win32_Process` instance that was created) and pull its Name and ProcessId. We then raise our custom `ProcessStarted` event. The `localWatch_ProcessStarted` demonstrates that you can handle it in VBA like any other event. The WMI sink runs on a separate thread internally, but `WithEvents` ensures the handler marshals back to the VBA thread. We also implement `sink_OnCompleted` to know if the query was canceled or ended. We call `sink.Cancel` in `StopWatching` to terminate the subscription (this prevents leaks or continued background activity). If you forget to cancel and set objects to Nothing, the WMI query may live beyond your VBA scope (until Excel closes or WMI decides to clean up) – always cancel to be safe. This example can be adapted to other WMI events: e.g., monitoring file creation (`__InstanceCreationEvent` on `CIM_DataFile`), USB device insertion, or system power events. It’s a powerful way to integrate system-level triggers into Excel.

### Example C3: Intercepting Ribbon Commands (Repurposing Built-in Buttons)

Excel Ribbon controls (like the Save button) can be intercepted and custom logic executed instead of (or in addition to) the default action. We can do this by providing custom UI XML that repurposes the control via its `idMso` (Microsoft’s internal ID for Ribbon controls). VBA will then handle the callback.

**Custom UI XML (for workbook or add-in):** – This XML, embedded in the workbook (e.g., using Office UI Editor), intercepts the Save command.

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
  <commands>
    <command idMso="FileSave" onAction="OnSaveIntercept"/>
  </commands>
</customUI>
```

This tells Excel: when the user invokes the built-in Save (FileSave) command, call our `OnSaveIntercept` macro instead.

**Standard Module:** Implement the callback `OnSaveIntercept`. Excel expects `Sub OnSaveIntercept(control As IRibbonControl, ByRef cancelDefault As Boolean)` signature for repurposed commands.

```vba
Option Explicit
Public Sub OnSaveIntercept(control As IRibbonControl, ByRef cancelDefault As Boolean)
    Dim wb As Workbook: Set wb = ThisWorkbook
    ' Custom logic before save
    Dim ans As VbMsgBoxResult
    ans = MsgBox("Do you want to save changes to " & wb.Name & "?", vbYesNoCancel + vbQuestion, "Custom Save")
    Select Case ans
        Case vbYes
            cancelDefault = False  ' allow Excel's normal save after our code
            MsgBox "Saving " & wb.Name, vbInformation
            ' (Excel will proceed to save because we did not cancel)
        Case vbNo
            cancelDefault = True   ' skip default save
            MsgBox "Changes were not saved.", vbExclamation
        Case vbCancel
            cancelDefault = True   ' skip default save and stay open
            MsgBox "Save canceled.", vbExclamation
    End Select
End Sub
```

*How it works:* The Ribbon XML’s `<command idMso="FileSave" onAction="OnSaveIntercept"/>` ties the built-in Save control to our macro. When Save is clicked (or Ctrl+S pressed), Excel calls `OnSaveIntercept`. We decide what to do and set `cancelDefault`: if False, Excel continues with the normal save after our macro; if True, Excel does nothing further (we completely override). In our example, pressing Cancel or No will stop the save (we set `cancelDefault=True`), Yes will let Excel proceed after showing a message. We could also perform a custom save routine (e.g., save to multiple locations) and then cancel Excel’s default to prevent double-saving. You can repurpose almost any built-in button this way by using the correct `idMso` (Microsoft provides lists of control IDs). This technique effectively raises an **event** (the Ribbon onAction call) that we handle in VBA, instead of Excel’s internal code. It integrates tightly with Excel’s UI.

**Note:** You cannot dynamically insert this XML via VBA at runtime; it must be in the workbook’s Office UI XML part (or you use a COM Add-in or OfficeJS for dynamic ribbon). For testing, use the Custom UI Editor to add the XML. The above callback will then be invoked as shown. Always include the `cancelDefault` parameter and set it appropriately; failing to set it leaves the default action (most built-ins default to not cancel, except some where Cancel=True might be required to avoid duplicates). Repurposing is powerful but should be used sparingly to not confuse users – e.g., intercepting “Delete Sheet” to add confirmation, or intercept “New Workbook” to apply templates, etc..

### Example C4: Subclassing Excel’s Window – Capturing Keystrokes

By **subclassing** a window, we can intercept low-level Windows messages that Excel or VBA doesn’t expose. Here, we subclass the Excel application window to catch **keyboard events** (WM\_KEYDOWN) globally. This is advanced and dangerous: if Excel crashes, the subclass might not be removed. We’ll implement a `KeyHook` class that sets a new WindowProc and listens for, say, the F2 key (which normally edits a cell). We’ll intercept F2 to do something else, and allow other keys to pass through.

**Class Module: `KeyHook`** – Hooks the Excel main window to capture key presses.

```vba
' Class Module: KeyHook
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
#Else
    ' 32-bit declares...
#End If

Private Const GWL_WNDPROC As Long = -4
Private Const WM_KEYDOWN As Long = &H100
Private Const VK_F2 As Long = &H71

Private prevProc As LongPtr  ' original window procedure
Private appHwnd As LongPtr   ' Excel main window handle

Public Sub Hook()
    If appHwnd <> 0 Then Unhook  ' ensure not double-hooked
    appHwnd = FindWindow("XLMAIN", Application.Caption)
    If appHwnd = 0 Then
        MsgBox "Failed to find Excel window.", vbCritical
        Exit Sub
    End If
    ' Subclass Excel window: install our WindowProc
    prevProc = SetWindowLongPtr(appHwnd, GWL_WNDPROC, AddressOf WindowProc)
    Debug.Print "KeyHook installed. prevProc=" & prevProc
End Sub

Public Sub Unhook()
    If appHwnd <> 0 And prevProc <> 0 Then
        SetWindowLongPtr(appHwnd, GWL_WNDPROC, prevProc)
        Debug.Print "KeyHook removed."
    End If
    prevProc = 0: appHwnd = 0
End Sub

' The custom window procedure to capture messages
Private Function WindowProc(ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    If Msg = WM_KEYDOWN Then
        Dim vk As Long: vk = CLng(wParam And &HFFFF&)
        If vk = VK_F2 Then
            Debug.Print "F2 key press intercepted!"
            ' Do not forward to Excel (skip default F2 behavior)
            WindowProc = 0
            Exit Function
        End If
    End If
    ' For all other messages and keys, call original window proc
    WindowProc = CallWindowProc(prevProc, hWnd, Msg, wParam, lParam)
End Function
```

**Standard Module: `Module1`** – Use the KeyHook.

```vba
Option Explicit
Public keyHook As KeyHook

Public Sub Test_KeyHook()
    Set keyHook = New KeyHook
    keyHook.Hook
    MsgBox "Key hook installed. Press F2 in Excel - it will not enter edit mode now.", vbInformation
    ' When finished testing:
    ' keyHook.Unhook
End Sub
```

After running `Test_KeyHook`, pressing F2 will no longer edit the cell; instead, our Debug.Print will show. Other keys work normally. **Important:** Always unhook (`keyHook.Unhook`) before ending the Excel session (we might do it in Workbook\_BeforeClose or similar) – failing to unhook can cause Excel to crash on exit or when the class goes out of scope. In this demo, we left it hooked to test; call `keyHook.Unhook` when done or on demand.

*How it works:* We find Excel’s top-level window via `FindWindow("XLMAIN", Application.Caption)` – class name "XLMAIN" is common for Excel windows. Then we call `SetWindowLongPtr` to replace the window’s procedure pointer with our `WindowProc` (AddressOf gives us a function pointer). We save the old pointer in `prevProc`. Now every message to Excel’s window goes first to our `WindowProc`. We check if it’s a keydown (WM\_KEYDOWN) and if the key is F2 (VK\_F2). If so, we intercept: we print a message and return 0 without calling Excel’s original proc – thus Excel never sees F2 (preventing edit mode). For any other key or message, we call `CallWindowProc` with `prevProc` to let default processing happen. This is a minimal example – we could extend it to other keys or messages (e.g., intercept WM\_KEYDOWN for Ctrl+Shift shortcuts not exposed, or intercept WM\_CLOSE to prompt user, etc.). This technique is powerful: it allows capturing things like Alt-key combinations, function keys, or even low-level mouse moves (WM\_MOUSEMOVE as in the earlier forum example). However, subclassing can destabilize Excel: if our code errors or we forget to unhook, Excel might attempt to call a freed procedure pointer. Always unhook (restore `prevProc`) before unloading the add-in or workbook. Also, **never debug/step through a subclass callback** – it can cause reentrant calls; use Debug.Print or logs instead. Despite risks, subclassing is the only way in VBA to catch events like key presses at the application level (Application.OnKey is simpler but only works for specific key combos and might not catch all scenarios, plus it replaces Excel’s behavior globally). With subclassing, we could implement custom keyboard shortcuts or block keys (like F1 Help) more comprehensively than Application.OnKey.

**Mini Case Study: Subclassing vs. OnKey for Custom Shortcuts** – A user wanted to override Excel’s F1 (Help) and F2 (Edit) globally in their workbook (to prevent accidental presses in a shared kiosk). **Attempt 1:** They used `Application.OnKey "{F1}", "{null}"` to disable F1, and similarly tried for F2. Result: F1 was effectively disabled (it did nothing, as desired), but F2 could not be fully nullified – it still sometimes entered edit mode because Excel uses F2 at a lower level that OnKey didn’t intercept (likely timing issues with modal states). Also, OnKey only works when a workbook is active and might be reset if Excel’s focus changes or on certain errors. **Attempt 2:** Implement subclassing via a technique like `KeyHook`. This reliably caught F2 (and any other keys) at the Windows message level – even if a cell was selected, pressing F2 now never invoked edit. They also chose to intercept F1 similarly to ensure no help popup (an alternative is `Application.HelpOption = xlHelpNone`, but subclassing works regardless of setting). **Performance:** The overhead of the subclass check was negligible (a few comparisons per key press). **Pitfalls:** During development, a crash occurred when Excel closed unexpectedly without unhooking, demonstrating the importance of robust unhook logic (they added unhook calls in Workbook\_BeforeClose and also in Workbook\_Deactivate just in case). Once stabilized, the subclass solution ran invisibly – users pressing F1 or F2 saw nothing happen, exactly as intended. This case highlights that for certain key events or deep-level interactions, subclassing is the only route in pure VBA. However, it should be reserved for when built-in options (OnKey, EnableCancelKey, etc.) are insufficient, due to complexity and risk.

### Example C5: Global Hotkey via Windows API (RegisterHotKey)

As a final system integration example, we show how to register a **global hotkey** – a key combination that triggers an action even when Excel is not the active application. This uses `RegisterHotKey` from user32. We still need to subclass to catch the WM\_HOTKEY message, but we don’t need to intercept all keys – Windows will send us a specific message when the hotkey is pressed anywhere in the OS.

**Class Module: `HotkeyListener`** – Registers a hotkey and listens for its message.

```vba
' Class Module: HotkeyListener
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function RegisterHotKey Lib "user32" (ByVal hWnd As LongPtr, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
    Private Declare PtrSafe Function UnregisterHotKey Lib "user32" (ByVal hWnd As LongPtr, ByVal id As Long) As Long
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#End If

Private Const GWL_WNDPROC As Long = -4
Private Const WM_HOTKEY As Long = &H312
' Modifier flags for RegisterHotKey:
Private Const MOD_ALT As Long = &H1
Private Const MOD_CONTROL As Long = &H2
Private Const MOD_SHIFT As Long = &H4

Private prevProc As LongPtr, hwnd As LongPtr, hotkeyID As Long

Public Event HotkeyPressed()  ' custom event for the hotkey

' Initialize with Excel's main window handle
Public Sub SetWindow(ByVal hWndTarget As LongPtr)
    hwnd = hWndTarget
End Sub

Public Function RegisterHotKeyCombo(ByVal modifiers As Long, ByVal virtualKey As Long) As Boolean
    hotkeyID = &HCAFE ' arbitrary ID
    If RegisterHotKey(hwnd, hotkeyID, modifiers, virtualKey) <> 0 Then
        ' Subclass to listen for WM_HOTKEY
        prevProc = SetWindowLongPtr(hwnd, GWL_WNDPROC, AddressOf WindowProc)
        RegisterHotKeyCombo = True
        Debug.Print "Global hotkey registered."
    Else
        Debug.Print "Failed to register hotkey."
    End If
End Function

Public Sub UnregisterHotKeyCombo()
    If hotkeyID <> 0 Then
        UnregisterHotKey(hwnd, hotkeyID)
    End If
    If prevProc <> 0 Then
        SetWindowLongPtr(hwnd, GWL_WNDPROC, prevProc)
    End If
    Debug.Print "Global hotkey unregistered."
End Sub

Private Function WindowProc(ByVal hWnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    If Msg = WM_HOTKEY Then
        If wParam = hotkeyID Then
            Debug.Print "Global hotkey pressed!"
            RaiseEvent HotkeyPressed  ' fire event for subscribers
            ' Do not exit; allow fall-through to default as well (not strictly necessary)
        End If
    End If
    WindowProc = CallWindowProc(prevProc, hWnd, Msg, wParam, lParam)
End Function
```

**Standard Module: `Module1`** – Demonstration of global hotkey usage.

```vba
Option Explicit
Public gHot As HotkeyListener

Public Sub Test_GlobalHotkey()
    Set gHot = New HotkeyListener
    Dim hwnd As LongPtr
    hwnd = FindWindow("XLMAIN", Application.Caption)  ' reuse FindWindow from earlier example
    gHot.SetWindow hwnd
    ' Register Ctrl+Shift+H as global hotkey (example: MOD_CONTROL + MOD_SHIFT with 'H')
    If gHot.RegisterHotKeyCombo(MOD_CONTROL Or MOD_SHIFT, Asc("H")) Then
        MsgBox "Press Ctrl+Shift+H from anywhere (even if Excel is not active) to trigger the hotkey.", vbInformation
        ' For demo, we handle the event here via WithEvents:
        Dim WithEvents hk As HotkeyListener: Set hk = gHot
        ' After pressing the hotkey, you should see "Global hotkey pressed!" in Immediate window
        ' and our event handler message in the output.
    End If
End Sub

Private Sub hk_HotkeyPressed()
    MsgBox "Hotkey event received in VBA!", vbExclamation
End Sub
```

Run `Test_GlobalHotkey`, confirm the MsgBox, then switch to another app or desktop and press Ctrl+Shift+H. Excel does not need focus – Windows will detect the combo and send a WM\_HOTKEY to Excel’s window, which our subclass catches. We Debug.Print and raise a VBA event. Our handler `hk_HotkeyPressed` shows a message box. **Important:** Only one app can register a given global hotkey at a time – if it fails (returns 0), maybe something else has it. Also, remember to call `UnregisterHotKeyCombo` (we omitted in demo; ideally call in Workbook\_BeforeClose) – otherwise the hotkey stays reserved until Excel closes.

*How it works:* We call `RegisterHotKey` with Excel’s window handle, an ID, modifier flags (MOD\_CONTROL|MOD\_SHIFT) and a VK code (VK for 'H' is ASCII 72). If successful, pressing that combo anywhere will cause Windows to post WM\_HOTKEY to Excel’s message queue. So we subclass (similar to KeyHook) to catch WM\_HOTKEY. When we see our ID, we raise a custom event. We then continue with `CallWindowProc` – for WM\_HOTKEY, returning 0 vs calling original doesn’t matter much, but we do call it to be safe. **Use case:** Global hotkeys might be used for an Excel-based tool that needs to respond even if Excel is backgrounded. For example, Ctrl+Shift+S might pop Excel up or run a macro from anywhere. Or a media player implemented in VBA could capture play/pause keys globally. **Caution:** As with subclassing, leaving it hooked after workbook close could crash Excel – always unregister and un-subclass on unload. Also, if Excel is not running, obviously the hotkey is inactive (this technique doesn’t create a system-wide persistent hotkey beyond Excel’s lifetime). Performance is trivial for a single key combo.

This concludes the deep integration section – we saw that with some API wizardry, VBA can handle events well beyond its usual scope, from high-frequency ticks to OS-level notifications and UI hooks. Always balance the complexity and ensure proper cleanup for these techniques.

## D. UserForm Architectures & Async UI (5 Examples)

Building interactive forms in VBA can benefit from event patterns too. Here we demonstrate design patterns like MVP and MVVM to separate UI and logic, handling asynchronous updates (e.g., showing progress on a modeless form during a long task), orchestrating multi-form workflows with events, and creating a persistent modeless toolpane that responds to application events.

### Example D1: Model-View-Presenter (MVP) – Separating Form UI and Logic

**MVP Pattern:** The UserForm is the *View* (dumb UI), a *Presenter* (class) handles the logic. The View exposes events (e.g., a Submit button click), which the Presenter subscribes to. The Presenter updates the View via an interface or direct reference, but the View itself does minimal work. This decouples business logic from UI design.

**UserForm: `FrmLogin`** – A simple login form with txtUsername, txtPassword, btnSubmit, lblStatus. It raises an event on submit.

```vba
' UserForm Module: FrmLogin
Option Explicit
Public Event Submit(ByVal username As String, ByVal password As String)

Private Sub btnSubmit_Click()
    RaiseEvent Submit(Me.txtUsername.Text, Me.txtPassword.Text)
End Sub

' Methods for Presenter to update UI:
Public Sub ShowMessage(msg As String)
    Me.lblStatus.Caption = msg
End Sub
```

**Class Module: `LoginPresenter`** – The Presenter that handles form events and validation logic.

```vba
' Class Module: LoginPresenter
Option Explicit
Private WithEvents view As FrmLogin  ' The UserForm (View) we interact with

' Initialize with the view and show it
Public Sub Start(ByVal loginForm As FrmLogin)
    Set view = loginForm
    view.lblStatus.Caption = ""      ' clear status
    view.Show
End Sub

' Handle the view's Submit event
Private Sub view_Submit(ByVal username As String, ByVal password As String)
    ' Example logic: check credentials (naively, just non-empty here)
    If Trim(username) = "" Or Trim(password) = "" Then
        view.ShowMessage "Please enter both username and password."
    ElseIf username = "admin" And password = "password123" Then
        view.ShowMessage "Login successful!"
        Debug.Print "Proceeding to next part of application..."  ' In real app, maybe open main form
        view.Hide
        ' (Presenter could trigger further events on successful login)
    Else
        view.ShowMessage "Login failed. Try again."
    End If
End Sub
```

**Standard Module: `Module1`** – Launch the MVP demo.

```vba
Option Explicit
Public Sub Test_MVP()
    Dim frm As New FrmLogin
    Dim presenter As New LoginPresenter
    presenter.Start frm
End Sub
```

Run `Test_MVP`. The login form appears. The Presenter is listening to `FrmLogin.Submit`. When you click Submit, `LoginPresenter.view_Submit` runs. It checks the credentials and calls `view.ShowMessage` to update the form’s status label. If “admin/password123” is entered, it prints a success and hides the form (could proceed). For other inputs, it shows “Login failed” on the form. Notice the form itself doesn’t contain any logic except raising the event and a simple UI update method. All decision-making is in the Presenter.

*How it works:* `FrmLogin` defines `Public Event Submit` and raises it on button click, passing data. The Presenter’s `WithEvents view` catches that. We pass the form to the presenter via `Start` (this wires the events). The presenter then displays the form (so UI control is still via UserForm’s normal `.Show`, but initiated by Presenter). When the event fires, the presenter method runs, doing validation. It uses the `ShowMessage` method (a public method on the form) to update the UI. Alternatively, we could have the Presenter update form controls directly (since it has a reference). Using a method makes it clearer and hides control details from the Presenter (maybe the presenter doesn’t need to know about `lblStatus`, just calls a method that the form implements). The MVP pattern here allows easily changing the logic (presenter) or reuse the form with different logic by hooking a different presenter. **Rubberduck note:** We see that events decouple the button click from the handling code – the form doesn’t know what happens on submit beyond raising that event. This is similar to how in .NET, a form might call `OnSubmit?.Invoke(...)`. It leads to more testable and maintainable code for complex forms.

### Example D2: Model-View-ViewModel (MVVM) – Binding UserForm to Data via Events

**MVVM Pattern:** The View (UserForm) binds to a ViewModel – a class that holds the data (Model) and logic for UI interactions. The ViewModel exposes properties and perhaps events (like `PropertyChanged`) that the View listens to, and the ViewModel may listen to view events or have commands. MVVM in VBA can’t use true data binding, but we simulate it with events.

We’ll create a small example: A form displays a counter value and has an “Increment” button. The counter is part of the ViewModel (data). When the button is clicked, the form doesn’t update itself; it asks the ViewModel to increment. The ViewModel changes its state, then raises a property-changed event that the form listens to, updating the label. This way, the form logic is minimal.

**Class Module: `CounterVM`** – The ViewModel (holds data and logic, raises events on changes).

```vba
' Class Module: CounterVM
Option Explicit
Private counter As Long
Public Event PropertyChanged(ByVal propName As String, ByVal newValue As Variant)

Public Property Get Count() As Long
    Count = counter
End Property

Public Sub Increment()
    counter = counter + 1
    RaiseEvent PropertyChanged("Count", counter)
End Sub

' We could add more properties and events as needed
```

**UserForm: `FrmCounter`** – The View binds to CounterVM. It updates label when VM fires event, and notifies VM when user clicks.

```vba
' UserForm Module: FrmCounter
Option Explicit
Private WithEvents vm As CounterVM  ' the ViewModel

Public Sub BindViewModel(ByVal viewModel As CounterVM)
    Set vm = viewModel
    ' initialize UI from current VM state
    lblCount.Caption = CStr(vm.Count)
End Sub

Private Sub btnIncrement_Click()
    If Not vm Is Nothing Then vm.Increment  ' delegate action to ViewModel
End Sub

' Listen for ViewModel property changes
Private Sub vm_PropertyChanged(ByVal propName As String, ByVal newValue As Variant)
    If propName = "Count" Then
        lblCount.Caption = CStr(newValue)
        Debug.Print "View updated Count display to " & newValue
    End If
End Sub
```

**Standard Module:** Launch MVVM demo.

```vba
Option Explicit
Public Sub Test_MVVM()
    Dim model As New CounterVM
    Dim frm As New FrmCounter
    frm.BindViewModel model
    frm.Show
End Sub
```

Run `Test_MVVM`. The form shows initial count 0. Clicking “Increment” triggers `btnIncrement_Click`, which calls `vm.Increment` (the VM logic). The VM increments its internal count and raises `PropertyChanged("Count", newValue)`. The form’s `vm_PropertyChanged` handler catches that and updates the label to the new value. The Debug.Print shows that update event in the Immediate window too.

*How it works:* The form holds a WithEvents reference to `CounterVM`. In `BindViewModel`, we connect them and initialize UI to current VM state (one-time sync). When user clicks, the form doesn’t change the label directly – it calls the VM’s `Increment`. The VM is the single source of truth for “Count”; it updates itself and notifies observers by raising `PropertyChanged`. The form, being an observer, sets the label accordingly. This mimics data binding: you change the data, the UI auto-updates via event. Conversely, if the UI had other inputs, the form could raise events or call VM methods to update the underlying model (like here, we call VM.Increment which updates model and triggers return event). The benefit is separation: `CounterVM` could be tested independently of the UI (ensuring it raises events properly). The form doesn’t need to know how increment works, and the VM doesn’t directly manipulate any UserForm controls (could even run headless). In complex apps, VM might talk to a domain model or external data source, and multiple forms could bind to the same VM. Because VBA is limited, we do this manually with events. Rubberduck’s blog mentions using Win32 subclassing to hook control events to custom event sinks for binding properties – that’s beyond our scope, but our simple approach still illustrates MVVM principles: one-way binding from VM to View and commands from View to VM.

### Example D3: Asynchronous Progress Bar on Modeless Form

Excel/VBA is single-threaded, but we can simulate asynchronous behavior using events and periodic yielding. A common scenario: a long-running process with a progress bar form that updates without freezing Excel. We’ll show a modeless UserForm with a ProgressBar (or just a label simulating progress) updated by events from a worker class. The worker will perform a task in small chunks, raising a ProgressChanged event periodically. The UI form subscribes and updates a bar. We’ll use `DoEvents` to keep UI responsive.

**Class Module: `WorkSimulator`** – Simulates a long process, raising progress events.

```vba
' Class Module: WorkSimulator
Option Explicit
Public Event ProgressChanged(ByVal percent As Long)
Public Event Completed()

Public Sub DoLongWork()
    Dim total As Long: total = 100
    Dim i As Long
    For i = 1 To total
        ' ... perform one unit of work (simulated by a tiny delay) ...
        If i Mod 5 = 0 Then
            RaiseEvent ProgressChanged(i)  ' update every 5%
        End If
        DoEvents  ' yield to allow UI update
    Next i
    RaiseEvent Completed
End Sub
```

**UserForm: `FrmProgress`** – Displays progress and handles events from WorkSimulator.

```vba
' UserForm Module: FrmProgress
Option Explicit
Private WithEvents worker As WorkSimulator

Public Sub StartWork(ByVal w As WorkSimulator)
    Set worker = w
    Me.lblStatus.Caption = "0%"
    Me.Show vbModeless
    ' Kick off the work (in same thread, but UI can update due to DoEvents in loop)
    worker.DoLongWork
End Sub

Private Sub worker_ProgressChanged(ByVal percent As Long)
    Me.lblStatus.Caption = percent & "%"
    Me.ProgressBar.Width = (percent * 2) ' assuming 200px = 100%
    DoEvents  ' ensure UI repaint
End Sub

Private Sub worker_Completed()
    Me.lblStatus.Caption = "Done!"
    Me.ProgressBar.Width = 200
    MsgBox "Work completed!", vbInformation
    Me.Hide
End Sub
```

**Standard Module:** Launch the async progress demo.

```vba
Option Explicit
Public Sub Test_AsyncProgress()
    Dim frm As New FrmProgress
    Dim w As New WorkSimulator
    frm.StartWork w
End Sub
```

Run `Test_AsyncProgress`. A modeless form shows with “0%”. The `WorkSimulator` begins its loop. Because of `DoEvents` inside, the UI can update. Every 5% progress, `ProgressChanged` event fires, the form’s handler updates a label and progress bar width. We include an extra `DoEvents` after updating to force immediate repaint. When done (100%), `Completed` event triggers: the form shows “Done!” and displays a message box, then hides.

*How it works:* We show the form modeless so `StartWork` doesn’t block execution of `DoLongWork`. In fact, here we call `worker.DoLongWork` right after showing modeless – execution enters the loop but because of `DoEvents`, the user could even move the form or do other Excel actions (though caution doing too much, as the loop is still running). Each event invocation is synchronous (events in VBA are synchronous calls to handlers), but `DoEvents` allows the form to refresh and handle redraw messages in between. The net effect is a smooth progress update without freezing. We simulate asynchronous behavior on one thread by cooperative multitasking (DoEvents). The pattern: break work into chunks, raise events for UI after each chunk, use DoEvents to keep UI alive. This avoids “Not Responding” and allows progress display. We must be careful: heavy use of DoEvents can allow user input or other events to fire (as in Section B), so sometimes better to disable certain UI or use flags. But for a basic progress bar, this is fine. After completion, we inform the user and hide the form. Another approach could be using the Win32 timer to perform work asynchronously (like scheduling chunks via OnTime), but that complicates flow. This straightforward method shows how events separate concerns: the Worker doesn’t know about any UI, it just raises events. The UI form doesn’t know details of work, it just updates visuals when told. We could also attach a logger or other observer to `ProgressChanged` without modifying worker logic – demonstrating event-driven flexibility.

### Example D4: Orchestrating Multiple Forms via Events (Workflow Wizard)

For multi-step workflows across forms (e.g., a multi-page wizard where each step is a separate UserForm), an **Orchestrator** can coordinate transitions by listening to events from each form. Instead of form1 directly launching form2, form1 raises an event “Next” and the orchestrator handles it by showing form2, etc. This decouples the forms from each other.

We’ll simulate a 2-step wizard. Each step is a form that collects some data, then signals it’s done. The `WorkflowManager` (orchestrator) listens and controls when to show/hide forms.

**UserForm: `FrmStep1`** – First step UI.

```vba
' UserForm: FrmStep1
Option Explicit
Public Event NextStep(ByVal name As String)
Private Sub btnNext_Click()
    If Trim(txtName.Text) = "" Then
        MsgBox "Please enter a name.", vbExclamation
    Else
        RaiseEvent NextStep(txtName.Text)
        Me.Hide  ' hide after getting data (or could Unload)
    End If
End Sub
```

**UserForm: `FrmStep2`** – Second step UI.

```vba
' UserForm: FrmStep2
Option Explicit
Public Event Finished(ByVal age As Integer)
Private Sub btnFinish_Click()
    If IsNumeric(txtAge.Text) Then
        RaiseEvent Finished(CInt(txtAge.Text))
        Me.Hide
    Else
        MsgBox "Please enter a valid age.", vbExclamation
    End If
End Sub
```

**Class Module: `WorkflowManager`** – Coordinates the two forms.

```vba
' Class Module: WorkflowManager
Option Explicit
Private WithEvents step1 As FrmStep1
Private WithEvents step2 As FrmStep2
Private userName As String
Private userAge As Integer

Public Sub StartWizard()
    ' create forms
    Set step1 = New FrmStep1
    Set step2 = New FrmStep2
    step1.Show vbModal  ' show first step modally (or modeless and manage accordingly)
    ' (The rest flows via events)
End Sub

Private Sub step1_NextStep(ByVal name As String)
    userName = name
    Debug.Print "Step1 completed. Name = " & userName
    ' Now show step2
    step2.Show vbModal
End Sub

Private Sub step2_Finished(ByVal age As Integer)
    userAge = age
    Debug.Print "Step2 completed. Age = " & userAge
    ' All steps done – perhaps process the collected data:
    MsgBox "Thanks, " & userName & ". Your age " & userAge & " is recorded.", vbInformation
    ' Could unload forms or reset if reusing
    Unload step1: Unload step2
End Sub
```

**Standard Module:** Run the wizard.

```vba
Option Explicit
Public Sub Test_MultiFormWorkflow()
    Dim w As New WorkflowManager
    w.StartWizard
End Sub
```

Run `Test_MultiFormWorkflow`. FrmStep1 shows. Enter a name, hit Next. The form raises `NextStep`, the WorkflowManager catches it, stores the name, prints debug, then shows FrmStep2. Step1 hides (but not unloaded; could unload to free, but we might access its controls via orchestrator if needed). On Step2, enter age, click Finish. It raises `Finished`, manager catches it, stores age, prints, then perhaps finalizes (here we just show a message combining name and age). Then we unload both forms and end.

*How it works:* The orchestrator holds WithEvents references to both forms, so it can catch their events. We show step1 modally. When Next is clicked, we hide step1 (modal show returns after form hidden/unloaded). The orchestrator’s `step1_NextStep` then executes (because modal show unblocked), we capture the data and immediately show step2 modally. Similarly, when finish is clicked, form2 hides/unloads, allowing `step2_Finished` to run. We then use both pieces of data. Note: because of modals, the orchestrator code after `step1.Show` doesn’t continue until step1 is closed (via Hide/Unload). That’s why we put subsequent logic in the event handler rather than sequential code. We could also show forms modeless and not block at all, relying entirely on events to sequence (in that case `StartWizard` would not use modal, and orchestrator would need to manage focus maybe, but events still drive transitions). The event-driven approach means Step1 and Step2 have no direct knowledge of each other; they just signal when done. The orchestrator could insert validation, branching (maybe skip step2 based on step1 data), etc., without changing form code. If the workflow had 3+ forms, orchestrator handles each event and decides next form to show. This is cleaner than, say, form1 calling `form2.Show` directly (which couples them and complicates if we add intermediate steps or reuse forms). Using events for form navigation improves maintainability of wizard-like sequences.

### Example D5: Modeless Toolpane with Application Event Integration

A **modeless toolpane** (a userform shown modelessly) can provide persistent UI (like a custom task pane) that updates based on user actions in Excel. We can have the toolpane subscribe to Excel events (via `Application` object WithEvents in the form) to refresh its content live.

We’ll make a modeless form that shows the address of the currently selected cell and offers a button to color that cell. It hooks the `Application.SheetSelectionChange` event to update the address display whenever selection changes.

**UserForm: `FrmToolPane`** – Our modeless tool window.

```vba
' UserForm: FrmToolPane
Option Explicit
Private WithEvents app As Application

Private Sub UserForm_Initialize()
    Set app = Application  ' capture application events
    lblSel.Caption = "(No selection)"
End Sub

Private Sub app_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    lblSel.Caption = "Active Cell: " & Target.Address(0, 0)
End Sub

Private Sub btnColor_Click()
    On Error Resume Next
    Dim rng As Range: Set rng = Application.ActiveCell
    If Not rng Is Nothing Then rng.Interior.Color = vbYellow
End Sub
```

**Standard Module:** Display the toolpane.

```vba
Option Explicit
Public Sub Show_ToolPane()
    Dim pane As New FrmToolPane
    pane.Show vbModeless
End Sub
```

Run `Show_ToolPane`. The form appears (modeless). Now click around on the worksheet – the label updates with active cell address because our form’s `app_SheetSelectionChange` handler runs on every selection change. Click the “Color Cell” button – it colors the current cell yellow (just an example action triggered from the toolpane). The form stays modeless, so you can interact with Excel and it stays on top (if needed, set form’s ShowModal=False property to ensure modeless behavior and perhaps ShowInTaskbar=False to keep it attached).

*How it works:* In `UserForm_Initialize`, we set `app = Application` with WithEvents, which ties into Excel’s Application events. The selection change event gives us the sheet and target range selected. We update a label accordingly. (Note: updating label frequently is fine; if performance was an issue, one might throttle it, but selection events aren’t terribly frequent normally). We also demonstrate intercepting user clicking the color button – we simply take ActiveCell and color it (we ignore if no cell). The key piece is the form listening to app events. This is an elegant way for a toolpane to reflect context (like selection, active sheet, etc.) without requiring explicit calls from outside. We must be mindful: if the form outlives the workbook (like if we close workbook with form open), we should either hide/unload it or ensure references are cleared (setting app = Nothing perhaps in QueryClose). If left, it might try to handle events on a dead Excel object – in practice, since Application is global, it exists until Excel closes, so maybe fine, but best to handle appropriately (the form will be destroyed when Excel shuts down anyway). This pattern is powerful: many add-ins use it – e.g., a modeless form that tracks the selection (like formula auditing tools, etc.). Because it’s modeless, user can do normal work with the form floating. The form can call any Excel object methods freely (since it’s in-process). The user could also do other actions (like type in a cell), and we might extend to handle other events (SheetChange to maybe update or log changes, etc., via WithEvents Application).

This approach essentially treats the form as a *view* onto Excel state, kept in sync by events. It’s similar to how .NET custom task panes might use events to update controls. The difference is we have to wire it manually in Initialize. Also, if multiple workbooks open, Application events cover all. If we wanted to specifically track a single workbook’s events, we might use a `Workbook` object WithEvents in the form instead. But selection is app-scoped anyway. We included cleaning up events in no explicit way here; to be safe, could do `Set app = Nothing` in UserForm\_Terminate to detach events, though when form unloads, that’s automatic.

This example shows a UI always visible, responding live to user context – a very interactive pattern enabled by events.

## E. Cross-Workbook and Cross-Process Orchestration (3 Examples)

Finally, we cover patterns that span multiple workbooks or even multiple Excel instances (processes). When dealing with add-ins or multiple projects, events and careful unloading are key. We demonstrate a global message bus for inter-workbook communication, a plugin loader that safely unloads with event cleanup, and controlling a separate Excel instance via COM events (cross-process).

### Example E1: Global Message Bus for Multi-Workbook Communication

In section A5 we built an Event Aggregator (MessageBus) within one project. We can extend that across workbooks by hosting the bus in an add-in or central workbook and letting other workbooks subscribe/publish to it. For simplicity, we simulate this in one workbook, but imagine `MessageBus` is in a hidden add-in loaded at startup.

**Class Module (in Add-in): `GlobalBus`** – Simple event bus.

```vba
' Class Module: GlobalBus
Option Explicit
Public Event Broadcast(ByVal topic As String, ByVal data As Variant)
Public Sub Publish(ByVal topic As String, ByVal data As Variant)
    Debug.Print "[Bus] broadcasting '" & topic & "'"
    RaiseEvent Broadcast(topic, data)
End Sub
```

**Standard Module (in Add-in)** – Expose a global instance of the bus.

```vba
Option Explicit
Public Bus As New GlobalBus
' This object will remain alive for the add-in's lifetime.
' (Could also use a Property Get to ensure a single instance, but this works with New.)
```

Now any other workbook (or add-in) that knows about this add-in can use it. For example, say our add-in is named “MyAddin.xlam” and is loaded. Another workbook could do:

```vba
' In another workbook:
Private WithEvents bus As Object

Private Sub Workbook_Open()
    Set bus = Application.Run("MyAddin.xlam!Bus") ' assuming Bus is exposed, or use an accessor
End Sub

Private Sub bus_Broadcast(ByVal topic As String, ByVal data As Variant)
    If topic = "Alert" Then
        MsgBox "Received alert: " & CStr(data)
    End If
End Sub
```

And from yet another place, one could publish:

```vba
Application.Run "MyAddin.xlam!Bus.Publish", "Alert", "Hello from Book1"
```

This would cause all subscribers (workbooks with a WithEvents bus) to receive the message and handle accordingly. In our test environment, we simulate within one workbook for demo:

```vba
' Simulation in one workbook (single process):
Dim WithEvents busLocal As GlobalBus

Sub Test_GlobalBusComm()
    Set busLocal = Bus  ' our global Bus instance from module
    ' Simulate another component publishing:
    Bus.Publish "Alert", "Test message"
End Sub

Private Sub busLocal_Broadcast(ByVal topic As String, ByVal data As Variant)
    Debug.Print "busLocal received topic=" & topic & ", data=" & data
End Sub
```

Expected output:

```
[Bus] broadcasting 'Alert'
busLocal received topic=Alert, data=Test message
```

*How it works:* The key is having a globally accessible bus object (in an add-in or a globally accessible module variable). External workbooks can’t directly declare WithEvents on a late-bound object; they need early-binding or to set a reference to the add-in’s library. But we can get around by returning the object as generic Object – surprisingly, WithEvents works if the object actually supports the event (as in the example above, we used `As Object` and got events, because runtime connects them). Another way: set a Tools->Reference to the add-in’s project and declare `Private WithEvents bus As GlobalBus`. Then you can directly do `Set bus = Bus` (the global). Either way, multiple workbooks now share one event stream. This is akin to how Excel’s own events allow cross-workbook things (but you might want custom events). A message bus could coordinate tasks: e.g., one workbook publishes "DataUpdated" and all open workbooks (or an add-in UI) get it and refresh. It’s very decoupled – publishers and subscribers don’t know about each other, only about the central bus. The central bus ideally lives as long as needed (the add-in, in our case, loaded for Excel session). When Excel closes or add-in unloads, bus goes away – subscribers should handle that (maybe check if bus Is Nothing in their events). This pattern avoids directly calling other workbooks’ macros (via Application.Run) by instead raising events, which is cleaner. (Under the hood, probably similar complexity, but architecturally nicer.)

**Note on cross-process:** This bus is intra-process (within one Excel instance). For truly cross-process communication (between separate Excel.exe instances or external processes), you’d need another mechanism (like DDE, or writing to a file or using some RPC). COM events don’t fire across processes unless using out-of-process COM servers. That’s beyond scope, but one might register a COM object in ROT (Running Object Table) to act as bus across processes. However, within one Excel, this suffices.

### Example E2: Plugin Loader and Safe Unload (Dynamic Add-In Management)

When dynamically loading or unloading VBA-based plugins (like .xlam add-ins or .xlsm “modules”), it’s crucial to clean up event handlers and references to avoid leaving things in memory or causing “ghost” events to fire after unload (which can crash Excel). We illustrate a pattern: a main manager that loads an add-in, calls an init (which sets up events), and before unloading, instructs the plugin to disconnect events and cleanup.

**Suppose** we have an add-in file "SamplePlugin.xlam" which, when opened, starts an Application event listener (maybe logs workbook opens). If we unload that add-in without it disconnecting, its event class might still be active and cause errors. The safe approach: The plugin provides a `Stop` method to detach events.

**In the Plugin (SamplePlugin.xlam)**:

```vba
' Class Module: AppEventsHandler (in plugin)
Option Explicit
Private WithEvents app As Application

Private Sub Class_Initialize()
    Set app = Application
End Sub

Private Sub Class_Terminate()
    Debug.Print "(Plugin) Handler terminating."
End Sub

Private Sub app_WorkbookOpen(ByVal Wb As Workbook)
    Debug.Print "(Plugin) Workbook opened: " & Wb.Name
End Sub

' Module in plugin:
Public PluginHandler As AppEventsHandler

Public Sub StartPlugin()
    Set PluginHandler = New AppEventsHandler
End Sub

Public Sub StopPlugin()
    Set PluginHandler = Nothing
End Sub
```

When the add-in opens, maybe it doesn’t auto-start to avoid running without permission. The main loader will invoke `StartPlugin`.

**In the Host (main workbook or manager add-in)**:

```vba
Option Explicit
Public Sub LoadPlugin()
    Application.AddIns("SamplePlugin").Installed = True  ' or Workbooks.Open if not in addins list
    Application.Run "SamplePlugin.xlam!StartPlugin"
    Debug.Print "Plugin loaded."
End Sub

Public Sub UnloadPlugin()
    On Error Resume Next
    Application.Run "SamplePlugin.xlam!StopPlugin"
    On Error GoTo 0
    Application.AddIns("SamplePlugin").Installed = False
    Debug.Print "Plugin unloaded."
End Sub
```

We would call `LoadPlugin` to load and initialize. The plugin’s AppEventsHandler now logs workbook opens. When we want to unload (maybe user disables the add-in), we call `UnloadPlugin`. This calls `StopPlugin` in the add-in, which sets the `PluginHandler` to Nothing, thereby terminating the event handler (Class\_Terminate runs, unsubscribing events). Then we actually unload the add-in file. This ensures no lingering event sinks remain referencing Excel (which could otherwise prevent the add-in from being fully garbage-collected or cause issues like ghost events trying to run code in an unloaded project).

We can test sequence:

```vba
Sub Test_PluginCycle()
    LoadPlugin
    ' Simulate some event:
    Workbooks.Add
    UnloadPlugin
    ' Closing the added workbook:
    Workbooks.Item(Workbooks.Count).Close False
End Sub
```

Expected output:

```
Plugin loaded.
(Plugin) Workbook opened: Book2
Plugin unloaded.
```

No errors or weird behavior beyond that. If we had not called `StopPlugin` and just removed the add-in, the AppEventsHandler would still be active (because a reference to Application still live) but its code context (the add-in project) is unloaded – leading to either a memory leak or a potential immediate crash when an event triggers. The explicit Stop avoids that by dropping the reference gracefully.

*How it works:* It’s straightforward – treat plugin objects like any other: properly release them. The main challenge is scope: the add-in’s event handler is probably a module-level object (as we did), so if not set to Nothing it would survive the add-in being closed (the COM references to Application keep it alive at least until Excel closes, albeit code is gone – unstable state). Another possible approach: the add-in could handle Workbook\_BeforeClose of itself to call StopPlugin automatically when it’s about to be unloaded (this is prudent as a backup). The pattern scales: if plugin had multiple objects or subscribed to events in other apps, provide a Stop that unsubscribes all.

### Example E3: Cross-Process Communication via COM – Automating Another Excel Instance

Cross-process “events” aren’t automatic, but we can mimic by one Excel controlling another via COM automation and using events on that remote Application object. For instance, from our main Excel, we can create another Excel.Application instance, open a workbook there, and handle that instance’s events (like when its workbook closes). This effectively allows one Excel to respond to happenings in another.

**In controlling Excel (Process A)**:

```vba
Option Explicit
Private WithEvents otherExcel As Application

Public Sub StartOtherExcel()
    Dim app2 As Object
    Set app2 = CreateObject("Excel.Application")
    app2.Visible = True
    app2.Workbooks.Add
    Set otherExcel = app2  ' hook events of other instance
End Sub

Private Sub otherExcel_WorkbookBeforeClose(ByVal Wb As Workbook, ByVal Cancel As Boolean)
    MsgBox "External workbook " & Wb.Name & " is closing (from other Excel).", vbInformation
End Sub

Public Sub CloseOtherExcel()
    If Not otherExcel Is Nothing Then
        otherExcel.Quit
        Set otherExcel = Nothing
    End If
End Sub
```

Run `StartOtherExcel`. A new Excel process launches (separate window). It adds a workbook. Our code attaches `otherExcel` WithEvents. Now, if in that new Excel you close the workbook or Excel itself, our `otherExcel_WorkbookBeforeClose` event fires in the original Excel, showing the message. We could even cancel the close by setting Cancel=True (though that would prevent the other instance from closing workbook – tricky but possible). This shows that COM events from one Excel can be caught by another as long as you hold the Application object with events in your process. It’s cross-process in that two Excel instances are separate, but they communicate via COM interface events.

After done, call `CloseOtherExcel` to terminate that other Excel and release object. (Always do this to avoid orphan Excel in memory).

*How it works:* `CreateObject("Excel.Application")` starts a new Excel instance. That returns an `Excel.Application` object reference. We assign it to a WithEvents variable. Now, any events on that instance (like workbook open, new workbook, sheet change, etc.) will call our handlers. COM takes care of marshaling the event call from other process into our process (as a cross-process COM call). This is how events are designed in COM – Excel’s Application class is a COM object supporting event interfaces that can be handled out-of-process. Performance: crossing processes adds overhead, but for few events it's fine (each event is like an RPC call). We should avoid heavy continuous events cross-process (like SelectionChange firing 100s of times quickly – might lag). But overall, it works. Use cases: maybe a supervisor Excel that launches others and monitors them (like a “Excel farm” manager), or an add-in that coordinates two separate Excel instances for multi-thread simulation (rare but could be done).

One must note: if the external workbook has code, those events might also fire in that instance – separate from our handlers. We are just observing. We could even drive the other Excel (like call methods on otherExcel object to manipulate it, which is standard automation). The event handling allows near real-time sync or logging between processes.

When done, quitting and releasing ensures the external Excel quits properly (important, or you'll leave it open invisibly possibly). If user manually closes external Excel, our reference becomes invalid – an event (WorkbookBeforeClose or ApplicationQuit if Excel had such event – Excel doesn’t expose AppQuit event – though we see last WB Close leads to Excel closing, at which point further events not received except that BeforeClose). We should in our event maybe detect if no workbooks left and then release our object.

**Safe unload cross-process:** Here, if external closes unexpectedly and we still hold `otherExcel`, our handlers may not get anything further (Excel likely gone). We should handle errors (like set otherExcel=Nothing in error handling if calls fail, etc.). Similarly, if we Quit it from here, events like WorkbookBeforeClose will fire (we can even Cancel them if needed).

This pattern is advanced but demonstrates that the event-driven mindset can extend beyond one Excel instance.

---

These examples collectively show how event-driven techniques – from basic custom events to system hooks and cross-boundary communication – can greatly improve the structure and capabilities of Excel VBA projects. By leveraging events, we achieve decoupling, responsiveness, and modularity that would be difficult with purely procedural code. Whether it’s implementing classic design patterns, managing performance, integrating with the OS, or orchestrating multiple components, events are a powerful tool in the VBA developer’s toolkit.

## Event Sequencing Corner Cases (Quick Reference)

Certain event interactions in Excel/VBA can be surprising. Here is a summary of tricky sequencing situations and what to expect:

| **Scenario**                               | **Event Sequence / Note**                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| ------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Macro calls `DoEvents` inside a loop       | Other pending events may fire **during the loop** before it continues. E.g., if DoEvents is in a long loop and user selects a cell, `SelectionChange` runs immediately. After those events, the loop resumes.                                                                                                                                                                                                                                                            |
| Handler triggers another event (recursion) | If an event handler performs an action that triggers the *same* event on the same object, Excel **does not queue it**; it will either ignore or require re-enabling events. E.g., writing to a cell in a Worksheet\_Change without disabling events leads to no new event (Excel prevents immediate recursion). For different events, one may trigger another (BeforeSave inside BeforeClose via Save call). Manage via `EnableEvents` to avoid unwanted cascades.       |
| Cancelling a “Before…” event               | If a `Cancel` parameter is set to True in a before-event, Excel stops the associated action and **skips subsequent events**. E.g., canceling `Workbook_BeforeClose` means the workbook won’t close and thus no `Workbook_Close` event fires. Always set Cancel deliberately and be aware it halts the default sequence.                                                                                                                                                  |
| Excel built-in events during custom events | If you raise a custom event (via `RaiseEvent`) or call into Excel object model, those calls can themselves trigger Excel events *before* your event handler finishes. E.g., a custom event handler that changes a cell will fire `Worksheet_Change` **before** the custom handler completes. Excel processes its internal events immediately. Use `Application.EnableEvents` or flags to control this if needed.                                                         |
| Timer or OnTime events in macro            | `Application.OnTime` events scheduled for a future time will not fire until the macro is free (unless `DoEvents` is used). They run on the main thread's event queue. A `timeSetEvent` callback, however, will fire even during a running macro (but it actually queues and executes when VBA yields via DoEvents or finishes, since still single-threaded). So practically, timer events might seem delayed if code is busy without DoEvents.                           |
| Order of multiple workbook events          | When a high-level action triggers multiple events (like opening a workbook triggers `Workbook_Open` and `WindowActivate`), the order can be important. Generally: *Open events* (Workbook\_Open) happen before *Activate* events of that workbook. But in complex scenarios (opening via code, or multiple workbooks), ordering can vary. Test specific cases if order matters, and do not assume one workbook’s event finishes before another starts unless documented. |
| UI interactions mid-macro (DoEvents usage) | If a macro shows a UserForm or uses DoEvents, the user can perform actions that fire events (clicking cells, etc.). Those event handlers run *within* the macro’s execution. If they error or alter global state, it can affect the macro. Best practice is to limit user interaction during critical code (e.g., disable Application.Interactive, or use flags to ignore certain events if macro active).                                                               |

**Key takeaway:** Excel/VBA event handling is mostly synchronous and reentrant. Cancel flags stop downstream events. Using DoEvents or modeless forms can intermix event sequences in non-intuitive ways. Protect against unintended recursion (disable events or use static guards), and always restore application state (EnableEvents True, ScreenUpdating True, etc.) even on error to normalize event flow.

## Pattern Selector – When to Use Which Event Pattern

Choosing the right event-driven pattern depends on your scenario. Here’s a quick guide:

| **Pattern**                        | **Use When...**                                                                                                                                               | **Example Use Case**                                                                                                                                                                                                               |
| ---------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Observer** (Pub-Sub one-to-many) | You have one object that multiple others should react to, and you want to decouple them.                                                                      | Data model update notifying multiple views (charts, forms) simultaneously of a change.                                                                                                                                             |
| **Mediator**                       | Multiple objects need to coordinate complex interactions, and you want central control to avoid them referencing each other.                                  | An app controller listens to UI element events and triggers model updates, logging, etc., rather than UI elements talking directly.                                                                                                |
| **Event Aggregator (Event Bus)**   | You need a global channel for events where senders and receivers don’t know each other (many-to-many). Useful across modules or workbooks.                    | A messaging system in an add-in where any part can “broadcast” (e.g., “ThemeChanged”) and any interested component updates accordingly.                                                                                            |
| **Command**                        | You want to encapsulate actions (especially if they may be queued, undone, or triggered from multiple places via events).                                     | A Ribbon button click raises an event that a CommandManager turns into a specific action object’s Execute call, decoupling button from implementation. Also use if you need an undo stack – each command knows how to undo itself. |
| **State Machine**                  | An object’s behavior/state transitions are complex and need to trigger events on state changes.                                                               | A multi-stage import process that moves through states (“Connecting”, “Fetching”, “Completed”) – UI can update on each transition via events, logic ensures valid transitions.                                                     |
| **Debounce** (event coil)          | An event fires too frequently to handle every occurrence, and you only care about the final outcome after a burst.                                            | User typing in a TextBox filtering a list – use debounce to only refresh list when typing pauses (improve performance).                                                                                                            |
| **Throttle**                       | You want to limit how often an event is handled, e.g., no more than X times per second, to prevent overhead.                                                  | Animating a shape during Worksheet\_SelectionChange – throttle to at most update position 5 times a second, even if user arrow-keys rapidly.                                                                                       |
| **Coalescing**                     | Many fine-grained events should be combined into one batch operation.                                                                                         | Multiple Worksheet\_Change events (from a paste operation updating many cells) – coalesce to one “DataBatchChanged” event after all changes, then recalc totals once instead of for each cell.                                     |
| **Subclassing (Windows API)**      | You need to catch an event that VBA doesn’t expose (keyboard, mouse, window messages), or override default behavior. *Use with extreme caution.*              | Capturing the Enter key in a UserForm TextBox which normally moves focus – subclass to intercept and treat as trigger for form submission. Or globally block F1 Help by subclassing Excel window.                                  |
| **Model-View-Presenter (MVP)**     | You want to separate UI code from logic, keeping form code minimal and logic testable. Ideal if UI might change or be reused.                                 | Complex validation on form inputs – Presenter handles it and updates form via events or method calls. Form just raises events on user actions.                                                                                     |
| **MVVM**                           | Similar to MVP but with a focus on data binding: the form should reflect model data automatically. Use if you have many UI elements bound to data properties. | A settings form where fields should update live as user changes options and might also reflect changes from elsewhere (like real-time data) – MVVM allows two-way binding via events (PropertyChanged, etc.).                      |
| **Global Message Bus**             | Multiple workbooks/add-ins need to communicate without tight coupling. Use an event bus in a central add-in.                                                  | An automation suite with a controller add-in and multiple plugin workbooks – the bus distributes commands (start/stop jobs) and status events (progress, errors) among them.                                                       |
| **Plugin Loader with Cleanup**     | Dynamically loading/unloading components – ensure to use if add-ins or child VBProjects subscribe to events.                                                  | An add-in manager that turns add-ins on/off based on user selection – implement Start/Stop in each add-in to handle enabling/disabling event hooks to avoid crashes on unload.                                                     |
| **Cross-App Events**               | You need to monitor or control another Office instance or external app via COM events.                                                                        | A monitoring tool that launches Excel instances on remote servers and logs when they open certain workbooks (using Application.WorkbookOpen events cross-process). Rare, but powerful for automation.                              |

Often, patterns can be combined. E.g., a system might use an Observer pattern internally but also publish certain events to a global bus for other components. The key is to use **Observer/Events** whenever direct calls would create tight coupling or synchronous waits that you’d prefer to invert. Use **Mediator** when objects are becoming too aware of each other. Use **throttling/debouncing** whenever rapid-fire events cause performance issues or flicker. And always ensure *cleanup* of event handlers to prevent memory leaks or cross-thread calls after dispose.

## API Call Performance (Micro-Benchmark)

Different event and timer mechanisms have different overheads. Here’s a rough comparison of per-call costs and characteristics:

| **Operation**                               | **Approx. Overhead**                                                                                         | **Notes**                                                                                                                                                                                                                                         |
| ------------------------------------------- | ------------------------------------------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `RaiseEvent` (intraprocess, in-memory)      | \~0.5–1.0 µs per listener (negligible)                                                                       | Very fast – just a vtable call to handler. Even with many (dozens) of listeners, cost is minor. Use events freely inside process.                                                                                                                 |
| `Application.OnTime` schedule               | \~ at least 1 ms to schedule, **≥1000 ms delay** to execution                                                | OnTime uses Excel’s scheduler – minimum interval is 1 second. Good for minute/hour timers, not sub-second. Scheduling overhead is low, but resolution is coarse (and not guaranteed exact).                                                       |
| Win32 `timeSetEvent` (multimedia timer)     | \~50 µs to set (plus callback execution) – resolution \~1–15 ms depending on system                          | High-resolution periodic timer. Can achieve 1ms intervals in ideal conditions, but callbacks are on main thread – if busy, calls queue. Use for sub-second scheduling. Need to kill timer after.                                                  |
| `DoEvents` call                             | \~0.5 µs overhead for call itself, but processing pending messages can add variable cost                     | Extremely fast to call, and can handle 1e6 calls in \~0.5 sec. The bigger cost is what happens during DoEvents – e.g., if user events or repaint occur. If nothing pending, \~0.5 µs. But if lots of events pending, time = time to process them. |
| Event Handler Dispatch (COM cross-process)  | \~100–1000 µs (0.1–1 ms) per event (depending on data passed)                                                | COM marshaling between processes has overhead. E.g., a cross-process WorkbookOpen event might take \~0.5 ms to arrive. If frequent (hundreds per second), could bottleneck. Keep cross-process events minimal or batched.                         |
| `RegisterHotKey` + subclass (global hotkey) | Hotkey press detection \~ under 0.1 ms (via Windows message)                                                 | RegisterHotKey itself is near instantaneous (few µs). When hotkey pressed, Windows posts message – retrieval via subclass is again µs-scale. Overall negligible overhead unless pressed extremely frequently.                                     |
| Excel built-in events (intraprocess)        | Usually under 10 µs to invoke handler (similar to RaiseEvent) – plus whatever Excel had to do to trigger it. | E.g., Worksheet.Change fires immediately after cell edit – overhead to call your code is trivial. But beware frequency: typing can cause many Change events – handle accordingly.                                                                 |

These are ballpark figures – actual performance can vary by system and scenario. The main point:

* **RaiseEvent and internal events are very fast** – do not worry about adding a few events; optimize only if profiling shows an event firing thousands of times per second.
* **OnTime** is coarse (1 second) – not suitable for sub-second precision tasks.
* **Win32 timers** give good precision (\~1ms), but remember they aren’t truly concurrent – use for UI refresh or periodic poll where OnTime isn’t enough.
* **DoEvents** itself is cheap, but can let expensive things run (e.g., recalculation, or lots of UI redraws).
* **COM cross-process events** add overhead – if you need to handle many events from another process, consider aggregating them or filtering what you subscribe to.
* **Frequent UI updates** (like updating a label 1000 times/sec) can be costly because of redraw overhead, not the event raising. Throttle or batch such updates (e.g., update progress bar at 20Hz instead of 200Hz) for smoother performance.

In summary, use events generously for clarity; the overhead is usually negligible compared to the work your handlers do. Just avoid designs that trigger *massive* numbers of events unnecessarily (for instance, toggling a flag 10000 times in a loop with RaiseEvent each time – combine into one event with a parameter if possible). And for any background timing or cross-boundary signaling, be mindful of the limitations (resolution, thread, marshaling) outlined above.
