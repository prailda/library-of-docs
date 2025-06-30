Принято. Я подготовлю исчерпывающий набор правил в формате JSON, охватывающий все синтаксические и прикладные аспекты написания кода VBA, включая конфликты имён, объявление переменных, передачу параметров, работу с массивами, объектами, событиями, ошибками и другие ключевые проблемы.

Каждое правило будет строго соответствовать шаблону: включая `must`, `must_not`, примеры плохого и хорошего кода, пояснения причины и последствий, а также структуру `uid`, `severity`, `scope` и другие элементы. Группировку я адаптирую для удобства использования LLM, логически структурируя по уровням синтаксиса, передачи данных, работы с памятью и типами.

Я сообщу, когда первый набор правил будет готов.


````json
[
  {
    "uid": "syntax/name_conflicts/enum_procedure",
    "title": "Конфликт имен: Enum и Sub",
    "severity": "compile_error",
    "vba_scope": ["StandardModule"],
    "context": "В одном модуле VBA нельзя объявлять Enum и процедуру с одинаковым именем.",
    "problem": "Одинаковые имена перечисления (Enum) и процедуры (Sub/Function) в одной области видимости приводят к неоднозначности.",
    "cause": "При компиляции объект видит дублирование идентификатора (именованная константа и процедура имеют одинаковый идентификатор).",
    "consequence": "Ошибка компиляции (Ambiguous name detected или Duplicate procedure name).",
    "rules": {
      "must": [
        "Гарантировать уникальность имён перечислений/элементов и процедур в рамках одного модуля.",
        "Использовать разные префиксы или слова для процедуры и перечисления (например, суффикс Enum для перечислений, глаголы для процедур)."
      ],
      "must_not": [
        "Не использовать одинаковое имя для Enum и любой процедуры.",
        "Не полагаться на модульную квалификацию – область имён общая для всего проекта."
      ]
    },
    "examples": {
      "bad": "```vba\nPublic Enum AnimalType\n    Dog = 1\n    Cat = 2\nEnd Enum\n\nPublic Sub Dog()\n    MsgBox \"Woof\"\nEnd Sub\n\nPublic Sub AnimalType()\n    MsgBox \"AnimalType sub\"\nEnd Sub\n```",
      "good": "```vba\nPublic Enum AnimalTypeEnum\n    AnimalType_Dog = 1\n    AnimalType_Cat = 2\nEnd Enum\n\nPublic Sub Dog_ShowMessage()\n    MsgBox \"Woof\"\nEnd Sub\n\nPublic Sub ProcessAnimalType()\n    MsgBox \"AnimalType sub\"\nEnd Sub\n```"
    },
    "rationale": "Перечисления и процедуры должны иметь разные неймспейсы. Уникальность имён предотвращает ошибки компиляции и облегчает чтение кода.",
    "tags": ["naming", "enum", "procedure"]
  },
  {
    "uid": "syntax/name_conflicts/event_procedure",
    "title": "Конфликт имен: событие и процедура",
    "severity": "compile_error",
    "vba_scope": ["ClassModule", "UserForm", "Worksheet", "Workbook"],
    "context": "В класс-модулях, UserForm или модулях листов процедуры с именами событий имеют фиксированную сигнатуру.",
    "problem": "Если объявить процедуру с именем события (или изменить сигнатуру обработчика события), возникнет конфликт.",
    "cause": "Компилятор ожидает строгое соответствие имени и параметров события, любое другое совпадение считается попыткой объявить обработчик.",
    "consequence": "Ошибка компиляции: \"Procedure declaration does not match description of event or procedure with the same name\".",
    "rules": {
      "must": [
        "Использовать точную сигнатуру для автоматических обработчиков событий (никаких лишних параметров).",
        "Давать уникальные имена процедурам, не являющимся обработчиками событий."
      ],
      "must_not": [
        "Не дублировать имя одного и того же события двумя процедурами в одном модуле.",
        "Не изменять сигнатуру процедуры события (не добавлять/удалять параметры)."
      ]
    },
    "examples": {
      "bad": "```vba\n' Неправильно: добавлен лишний параметр в обработчик UserForm\nPrivate Sub UserForm_Initialize(ByVal InitFlag As Integer)\n    ' ... код ...\nEnd Sub\n```",
      "good": "```vba\n' Правильно: корректная сигнатура события без дополнительных параметров\nPrivate Sub UserForm_Initialize()\n    ' ... код инициализации формы ...\nEnd Sub\n```"
    },
    "rationale": "Имя процедуры события зарезервировано средой VBA. Нарушение сигнатуры или дублирование имени приводит к ошибке компиляции.",
    "tags": ["events", "signature", "naming"]
  },
  {
    "uid": "syntax/name_conflicts/property_procedure",
    "title": "Конфликт имен: Property и Sub/Function",
    "severity": "compile_error",
    "vba_scope": ["ClassModule", "Worksheet", "UserForm"],
    "context": "В классе или форме свойство реализуется через процедуры Get/Let/Set с одним именем, которое не может совпадать с именем другой процедуры.",
    "problem": "Одноимённое свойство и процедура приводят к конфликту – компилятор не сможет отличить их.",
    "cause": "При объявлении свойства (Property) создаётся имя, в которое нельзя в дальнейшем поместить процедуру с таким же именем.",
    "consequence": "Ошибка компиляции (Duplicate procedure name или Ambiguous name detected) – имя занято.",
    "rules": {
      "must": [
        "Обеспечивать уникальность имен свойств и методов в одном классе/форме.",
        "Если метод выполняет действие по концепции свойства, объединить логику или дать методу другое имя (обычно глагол)."
      ],
      "must_not": [
        "Не давать процедуре имя, совпадающее с уже существующим свойством или элементом управления.",
        "Не полагаться на регистр – VBA не различает имён регистром."
      ]
    },
    "examples": {
      "bad": "```vba\n' В классе MyClass\nPrivate mTotal As Long\n\nPublic Property Get Total() As Long\n    Total = mTotal\nEnd Property\n\nPublic Property Let Total(ByVal newValue As Long)\n    mTotal = newValue\nEnd Property\n\n' Конфликт: процедура с тем же именем\nPublic Sub Total()\n    mTotal = 0\nEnd Sub\n```",
      "good": "```vba\n' В классе MyClass - после исправления\nPrivate mTotal As Long\n\nPublic Property Get Total() As Long\n    Total = mTotal\nEnd Property\n\nPublic Property Let Total(ByVal newValue As Long)\n    mTotal = newValue\nEnd Property\n\n' Процедура получила уникальное имя\nPublic Sub ResetTotal()\n    mTotal = 0\nEnd Sub\n```"
    },
    "rationale": "Имя свойства уже зарезервировано за Get/Let/Set методами, поэтому добавлять процедуру с таким же именем невозможно.",
    "tags": ["naming", "property", "method"]
  },
  {
    "uid": "syntax/name_conflicts/enum_type",
    "title": "Конфликт имен: Enum и Type",
    "severity": "compile_error",
    "vba_scope": ["StandardModule"],
    "context": "Enum и пользовательские Type находятся в общей глобальной области имён VBA-проекта.",
    "problem": "Использование одинакового имени для перечисления (Enum) и пользовательского типа (Type) вызывает конфликт.",
    "cause": "В области проекта неясно, что именно представляет имя – enum или user-defined type.",
    "consequence": "Ошибка компиляции (Name conflicts) – код не компилируется.",
    "rules": {
      "must": [
        "Присваивать разное имя Enum и Type (например, добавлять суффикс или менять имя).",
        "Убедиться, что все публичные типы имеют уникальные имена."
      ],
      "must_not": [
        "Не объявлять Enum и Type с одинаковым именем в одном проекте.",
        "Не полагаться на специфику модуля – проблема на уровне проекта."
      ]
    },
    "examples": {
      "bad": "```vba\nPublic Enum ClientData\n    Individual = 1\n    Company = 2\nEnd Enum\n\nPublic Type ClientData\n    Name As String\n    ID   As Long\nEnd Type\n```",
      "good": "```vba\nPublic Enum ClientCategory\n    ClientIndividual = 1\n    ClientCompany = 2\nEnd Enum\n\nPublic Type ClientData\n    Name     As String\n    ID       As Long\n    Category As ClientCategory\nEnd Type\n```"
    },
    "rationale": "Все публичные объекты и типы данных должны иметь уникальные имена, чтобы избежать путаницы и ошибок компиляции.",
    "tags": ["naming", "enum", "type"]
  },
  {
    "uid": "syntax/keyword_conflict/variable",
    "title": "Зарезервированное слово как имя переменной",
    "severity": "compile_error",
    "vba_scope": ["StandardModule"],
    "context": "Нельзя использовать ключевые слова VBA как имена переменных.",
    "problem": "Попытка объявить переменную с именем, совпадающим с ключевым словом, приводит к синтаксической ошибке.",
    "cause": "Имя переменной совпадает с языковым зарезервированным словом.",
    "consequence": "Ошибка компиляции (Unexpected end of statement или синтаксическая ошибка).",
    "rules": {
      "must": [
        "Переименовать переменную, добавив префикс/суффикс так, чтобы она не совпадала с ключевым словом."
      ],
      "must_not": [
        "Не использовать ключевые слова (Next, Type, Exit и т.д.) в качестве идентификаторов."
      ]
    },
    "examples": {
      "bad": "```vba\nDim Next As Integer      ' Ошибка: Next – зарезервированное слово\n```",
      "good": "```vba\nDim NextItem As Integer  ' Правильно: изменено имя, конфликта нет\n```"
    },
    "rationale": "Ключевые слова несут специальный смысл в языке, их использование в качестве имён нарушает синтаксис.",
    "tags": ["keywords", "naming"]
  },
  {
    "uid": "syntax/keyword_conflict/procedure",
    "title": "Зарезервированное слово как имя процедуры",
    "severity": "compile_error",
    "vba_scope": ["StandardModule", "ClassModule"],
    "context": "Нельзя давать процедурам (Sub/Function) имена, совпадающие с ключевыми словами VBA.",
    "problem": "Имя процедуры совпадает с зарезервированным словом (например, Type), что нарушает синтаксис.",
    "cause": "Использование ключевого слова вместо нового имени процедуры.",
    "consequence": "Ошибка компиляции: идентификатор недопустим.",
    "rules": {
      "must": [
        "Переименовать процедуру, избегая использования ключевых слов."
      ],
      "must_not": [
        "Не называть Sub/Function ключевыми словами (Type, Next, End и т.д.)."
      ]
    },
    "examples": {
      "bad": "```vba\nSub Type()\n    ' ...\nEnd Sub\n```",
      "good": "```vba\nSub DoType()\n    ' ...\nEnd Sub\n```"
    },
    "rationale": "Ключевые слова имеют специальное синтаксическое значение, использование их нарушает правила парсинга.",
    "tags": ["keywords", "procedure"]
  },
  {
    "uid": "syntax/keyword_conflict/enum_element",
    "title": "Зарезервированное слово как имя элемента Enum",
    "severity": "compile_error",
    "vba_scope": ["StandardModule"],
    "context": "В Enum также нельзя называть элементы именами ключевых слов.",
    "problem": "Имя элемента Enum совпадает с ключевым словом (например, Error, Exit).",
    "cause": "При объявлении Enum элементу присвоено зарезервированное слово.",
    "consequence": "Ошибка компиляции: зарезервированное слово не может быть использовано.",
    "rules": {
      "must": [
        "Использовать альтернативные, нефункциональные имена для элементов Enum."
      ],
      "must_not": [
        "Не называть элементы Enum зарезервированными словами."
      ]
    },
    "examples": {
      "bad": "```vba\nPublic Enum MyEnum\n    [Error] = 1   ' Ошибка: Error – зарезервированное слово\n    [Exit] = 2    ' Ошибка: Exit – ключевое слово\nEnd Enum\n```",
      "good": "```vba\nPublic Enum MyEnum\n    ValueA = 1  ' Правильно: не ключевое слово\n    ValueB = 2\nEnd Enum\n```"
    },
    "rationale": "Зарезервированные слова нельзя использовать в контексте объявления, даже в квадратных скобках.",
    "tags": ["keywords", "enum"]
  },
  {
    "uid": "syntax/operators/other_language",
    "title": "Операторы других языков в VBA",
    "severity": "compile_error",
    "vba_scope": ["StandardModule"],
    "context": "Некоторые операторы из C/C++ (==, !=, &&, ||, ++ и т.д.) не поддерживаются в VBA.",
    "problem": "Использование недопустимых операторов приводит к ошибке компиляции.",
    "cause": "Перепутан синтаксис, ожидаемый в VBA.",
    "consequence": "Ошибка компиляции (Unexpected token или неверный синтаксис).",
    "rules": {
      "must": [
        "Использовать только операторы VBA: `=`, `<>`, `And`, `Or`. Для инкремента/декремента выполнять явное присваивание."
      ],
      "must_not": [
        "Не применять операторы `==`, `!=`, `&&`, `||`, `++`, `+=` и т.п. в VBA."
      ]
    },
    "examples": {
      "bad": "```vba\nIf x == 10 Then   ' Ошибка: оператор == не существует\nIf a != b Then    ' Ошибка: != не используется (в VBA - <>)\nIf flag && cond Then  ' Ошибка: && не поддерживается (используйте And)\nIf flag || cond Then  ' Ошибка: || не поддерживается (используйте Or)\ncount += 1          ' Ошибка: += не работает в VBA\ni++                 ' Ошибка: ++ недопустим\n```",
      "good": "```vba\nIf x = 10 Then\nIf a <> b Then\nIf flag And cond Then\nIf flag Or cond Then\ncount = count + 1\ni = i + 1\n```"
    },
    "rationale": "VBA использует другие обозначения. Неправильные операторы нарушают синтаксис и логику кода.",
    "tags": ["operators", "syntax"]
  },
  {
    "uid": "syntax/array_parameters/bound_in_declaration",
    "title": "Указание границ массива в параметрах",
    "severity": "compile_error",
    "vba_scope": ["StandardModule"],
    "context": "При объявлении параметра-массива нельзя указывать его размерность.",
    "problem": "Задание границ массива в объявлении процедуры запрещено.",
    "cause": "Неверный синтаксис параметра.",
    "consequence": "Ошибка компиляции (Sub or Function declaration not allowed here).",
    "rules": {
      "must": [
        "Объявлять параметр-массив без указания границ: `nums() As Type`."
      ],
      "must_not": [
        "Не указывать `(1 To N)` или другие границы в списке параметров."
      ]
    },
    "examples": {
      "bad": "```vba\nSub PrintNumbers(nums(1 To 10) As Integer)\n    ' ...\nEnd Sub\n```",
      "good": "```vba\nSub PrintNumbers(nums() As Integer)\n    ' ...\nEnd Sub\n```"
    },
    "rationale": "Параметр-массив должен быть объявлен как динамический, чтобы можно было передавать массивы различной длины.",
    "tags": ["parameters", "arrays"]
  },
  {
    "uid": "syntax/array_parameters/byval_for_array",
    "title": "Передача массива по значению (ByVal)",
    "severity": "compile_error",
    "vba_scope": ["StandardModule"],
    "context": "Массивы в VBA всегда передаются по ссылке (ByRef).",
    "problem": "Использование `ByVal` для параметра-массива недопустимо.",
    "cause": "Язык не поддерживает копирование массива при вызове.",
    "consequence": "Ошибка компиляции (Sub or Function declaration not allowed here).",
    "rules": {
      "must": [
        "Объявлять массив-параметр без модификатора (по умолчанию ByRef) или явно использовать `ByRef`."
      ],
      "must_not": [
        "Не указывать `ByVal` у параметра-массива."
      ]
    },
    "examples": {
      "bad": "```vba\nFunction SumArray(ByVal arr() As Long) As Long\n    ' ...\nEnd Function\n```",
      "good": "```vba\nFunction SumArray(ByRef arr() As Long) As Long\n    ' ...\nEnd Function\n```"
    },
    "rationale": "В VBA массивы передаются по ссылке. Попытка ByVal вызовет ошибку компиляции.",
    "tags": ["parameters", "arrays"]
  },
  {
    "uid": "syntax/set_misuse/missing_set_for_object",
    "title": "Пропущен Set при присвоении объекту",
    "severity": "compile_error",
    "vba_scope": ["StandardModule"],
    "context": "Для присваивания объектным переменным необходимо использовать `Set`.",
    "problem": "Отсутствие `Set` при присваивании объекта ведёт к ошибке типа.",
    "cause": "При попытке присвоить объект без `Set` VB воспринимает операцию как попытку присвоить значение по Value.",
    "consequence": "Ошибка компиляции или Type Mismatch на этапе выполнения.",
    "rules": {
      "must": [
        "При присвоении объектных переменных всегда использовать `Set`."
      ],
      "must_not": [
        "Не пропускать `Set` при работе с объектами."
      ]
    },
    "examples": {
      "bad": "```vba\nDim dict As Object\ndict = CreateObject(\"Scripting.Dictionary\")  ' Ошибка: нет Set\n```",
      "good": "```vba\nDim dict As Object\nSet dict = CreateObject(\"Scripting.Dictionary\")  ' Правильно: используем Set\n```"
    },
    "rationale": "Оператор Set необходим для присваивания ссылок на объекты. Без него привязывается значение (Value), что для объектов недопустимо.",
    "tags": ["objects", "set"]
  },
  {
    "uid": "syntax/set_misuse/unnecessary_set",
    "title": "Лишний Set при присвоении значения",
    "severity": "compile_error",
    "vba_scope": ["StandardModule"],
    "context": "Оператор `Set` используется только для объектных переменных.",
    "problem": "Использование `Set` для обычных типов (Integer, String и др.) вызывает ошибку.",
    "cause": "Попытка применить Set к переменной значимого типа.",
    "consequence": "Ошибка компиляции (Invalid use of property или аналогичная).",
    "rules": {
      "must": [
        "Просто присваивать значения без `Set` для скалярных переменных."
      ],
      "must_not": [
        "Не использовать `Set` при присваивании чисел, строк или других не-объектных типов."
      ]
    },
    "examples": {
      "bad": "```vba\nDim count As Integer\nSet count = 5   ' Ошибка: Set применяется только к объектам\n```",
      "good": "```vba\nDim count As Integer\ncount = 5       ' Правильно: обычное присваивание\n```"
    },
    "rationale": "`Set` не предназначен для примитивных типов; его использование приводит к синтаксической ошибке.",
    "tags": ["objects", "set"]
  },
  {
    "uid": "syntax/public/class_module",
    "title": "Неправильное использование Public в ClassModule",
    "severity": "compile_error",
    "vba_scope": ["ClassModule"],
    "context": "В класс-модулях `Public` работает иначе, чем в стандартных модулях.",
    "problem": "Нельзя объявлять `Public Const`, статические массивы или `Type` внутри класса.",
    "cause": "Язык не поддерживает эти конструкции как публичные в классах.",
    "consequence": "Ошибка компиляции: \"Constants, fixed-length strings, arrays... not allowed as Public members...\".",
    "rules": {
      "must": [
        "Если требуется глобальная константа или массив, объявлять их в стандартном модуле.",
        "Доступ к Public методам класса осуществлять через экземпляр класса."
      ],
      "must_not": [
        "Не объявлять внутри класса Public константы, фиксированные массивы или пользовательские типы.",
        "Не полагаться на то, что Public-члены класса становятся глобальными."
      ]
    },
    "examples": {
      "bad": "```vba\n' В ClassModule:\nPublic Const PI As Double = 3.14    ' Ошибка\nPublic arr(1 To 5) As String        ' Ошибка\nPublic Type Person                  ' Ошибка\n    Name As String\nEnd Type\n```",
      "good": "```vba\n' В ClassModule:\nPublic Value As Integer   ' Допустимо (Public примитивный тип)\n\n' В стандартном модуле:\nPublic Const PI As Double = 3.14  ' Правильно для глобальной константы\n```"
    },
    "rationale": "В класс-модулях можно объявлять Public только методы, свойства и примитивные поля. Другие элементы должны быть в стандартном модуле или приватными.",
    "tags": ["modifiers", "class"]
  },
  {
    "uid": "syntax/event_declaration/standard_module",
    "title": "Объявление Event в стандартном модуле",
    "severity": "compile_error",
    "vba_scope": ["StandardModule"],
    "context": "Объявлять события (`Event`) разрешается только в модулях классов или UserForm.",
    "problem": "Попытка объявить `Event` в обычном модуле вызовет синтаксическую ошибку.",
    "cause": "Оператор `Event` недопустим вне класса/формы.",
    "consequence": "Ошибка компиляции: синтаксическая ошибка при использовании `Event` в обычном модуле.",
    "rules": {
      "must": [
        "Объявлять события только в класс-модулях или модулях форм."
      ],
      "must_not": [
        "Не использовать `Event` в стандартных модулях."
      ]
    },
    "examples": {
      "bad": "```vba\n' В стандартном модуле Module1:\nPublic Event DataLoaded()   ' Ошибка: Event недопустим здесь\n```",
      "good": "```vba\n' В модуле класса ClassModule1:\nPublic Event DataLoaded()   ' Правильно: Event внутри класса\n```"
    },
    "rationale": "События являются частью объектной модели и могут объявляться только внутри классов.",
    "tags": ["events", "modules"]
  },
  {
    "uid": "syntax/interface/implements_outside_class",
    "title": "Implements вне класса",
    "severity": "compile_error",
    "vba_scope": ["StandardModule"],
    "context": "`Implements` можно использовать только в класс-модулях.",
    "problem": "Применение `Implements` в обычном модуле невозможно.",
    "cause": "Оператор `Implements` вне ClassModule.",
    "consequence": "Ошибка компиляции: \"Invalid outside Type block\" или аналогичная.",
    "rules": {
      "must": [
        "Использовать `Implements` только внутри ClassModule."
      ],
      "must_not": [
        "Не писать `Implements` в стандартных модулях."
      ]
    },
    "examples": {
      "bad": "```vba\n' В стандартном модуле Module1:\nImplements IMyInterface   ' Ошибка: недопустимо вне класса\n```",
      "good": "```vba\n' В модуле класса Class1:\nImplements IMyInterface   ' Правильно: реализация интерфейса в классе\n```"
    },
    "rationale": "Интерфейсы реализуются только в класс-модулях. В модуле общего назначения `Implements` недействителен.",
    "tags": ["interface"]
  },
  {
    "uid": "syntax/constants/declaration",
    "title": "Объявление Const: инициализация и выражения",
    "severity": "compile_error",
    "vba_scope": ["StandardModule"],
    "context": "Константу `Const` необходимо сразу инициализировать литералом или выражением, вычисляемым на этапе компиляции.",
    "problem": "Объявление константы без значения или с выражением, вычисляемым во время выполнения, недопустимо.",
    "cause": "Константа не была проинициализирована или использовалось не константное выражение (переменная, функция).",
    "consequence": "Ошибка компиляции: \"Требуется константное выражение\" или синтаксическая ошибка.",
    "rules": {
      "must": [
        "Задавать значение `Const` сразу при объявлении.",
        "Использовать только литералы, ранее определённые константы или арифметические выражения из них."
      ],
      "must_not": [
        "Не объявлять `Const` без значения.",
        "Не использовать переменные, функции или другие вычисления времени выполнения в инициализации."
      ]
    },
    "examples": {
      "bad": "```vba\nConst X As Integer          ' Ошибка: нет значения\nX = 5\nDim n As Integer: n = 10\nConst Y As Integer = n    ' Ошибка: n не известно\nConst Z As String = Now() ' Ошибка: Now() – не константное выражение\n```",
      "good": "```vba\nConst XMax As Integer = 5\nConst X2 As Integer = 2 * 2 + 1\nPrivate Const Msg As String = \"Hello\"\n```"
    },
    "rationale": "Константа должна быть известна на этапе компиляции. Это позволяет избежать непредсказуемого поведения и ошибок.",
    "tags": ["constants", "syntax"]
  },
  {
    "uid": "parameters/object_type",
    "title": "Необоснованное использование Object",
    "severity": "warning",
    "vba_scope": ["StandardModule"],
    "context": "Объявление параметра как `Object` без необходимости приводит к позднему связыванию.",
    "problem": "Отсутствие проверки типа на этапе компиляции – несовместимый объект вызовет ошибку во время выполнения.",
    "cause": "Указан слишком общий тип `Object` вместо конкретного.",
    "consequence": "Runtime error при несовместимом объекте; отсутствие IntelliSense и производительность ниже.",
    "rules": {
      "must": [
        "Указывать конкретный класс или интерфейс для параметров, если известен ожидаемый тип."
      ],
      "must_not": [
        "Не объявлять все объекты как Object без надобности."
      ]
    },
    "examples": {
      "bad": "```vba\nSub CloseBook(obj As Object)\n    obj.Close   ' Ошибка, если obj не Workbook-like объект\nEnd Sub\n```",
      "good": "```vba\nSub CloseBook(wb As Workbook)\n    wb.Close   ' Безопасно: тип проверяется компилятором\nEnd Sub\n```"
    },
    "rationale": "Явные типы дают раннее связывание, повышают производительность и надёжность.",
    "tags": ["parameters", "types"]
  },
  {
    "uid": "parameters/optional/missing_check",
    "title": "Необработанный Optional-параметр",
    "severity": "runtime_error",
    "vba_scope": ["StandardModule"],
    "context": "Optional-параметры типа Variant без значения получают состояние Missing.",
    "problem": "Использование Optional параметра без проверки приводит к ошибке при сравнении.",
    "cause": "Не выполнена проверка IsMissing перед использованием параметра.",
    "consequence": "Runtime error или логические ошибки.",
    "rules": {
      "must": [
        "Использовать `IsMissing` для Optional параметров Variant перед их использованием."
      ],
      "must_not": [
        "Не обращаться к Optional параметру без проверки его наличия."
      ]
    },
    "examples": {
      "bad": "```vba\nSub SendMessage(Optional msg As Variant)\n    If msg = \"OK\" Then Call ProcessOK\nEnd Sub\n```",
      "good": "```vba\nSub SendMessage(Optional msg As Variant)\n    If IsMissing(msg) Then Exit Sub\n    If msg = \"OK\" Then Call ProcessOK\nEnd Sub\n```"
    },
    "rationale": "Проверка Optional параметров предотвращает ошибки при попытке использовать несуществующее значение.",
    "tags": ["parameters", "optional"]
  },
  {
    "uid": "parameters/positional_order",
    "title": "Неправильный порядок позиционных аргументов",
    "severity": "runtime_error",
    "vba_scope": ["StandardModule"],
    "context": "При вызове процедуры с позиционными аргументами важен правильный порядок.",
    "problem": "Перестановка аргументов может привести к неправильному значению или ошибке преобразования типов.",
    "cause": "Позиционные аргументы переданы не в том порядке.",
    "consequence": "Логически неверные результаты или Type Mismatch.",
    "rules": {
      "must": [
        "Следовать объявленному порядку параметров в вызове.",
        "Или использовать именованные аргументы для гибкости порядка."
      ],
      "must_not": [
        "Не менять местами позиционные аргументы без явной необходимости."
      ]
    },
    "examples": {
      "bad": "```vba\nSub Connect(host As String, port As Integer)\nEnd Sub\n\nConnect 8080, \"Server1\"   ' Ошибка: порядок нарушен\n```",
      "good": "```vba\nConnect \"Server1\", 8080    ' host=\"Server1\", port=8080\n```"
    },
    "rationale": "VBA не предупреждает об ошибке, если типы совместимы, но значения будут перепутаны.",
    "tags": ["parameters", "order"]
  },
  {
    "uid": "parameters/named_usage",
    "title": "Неправильные именованные аргументы",
    "severity": "compile_error",
    "vba_scope": ["StandardModule"],
    "context": "Именованные аргументы должны точно соответствовать имени параметра и не дублироваться.",
    "problem": "Опечатка в имени или повторный указанный аргумент вызывает ошибку.",
    "cause": "Неправильно написано имя параметра или задано дважды.",
    "consequence": "Ошибка компиляции: именованный параметр не найден или дублирован.",
    "rules": {
      "must": [
        "Писать имена параметров точно, учитывая регистр и отсутствие опечаток."
      ],
      "must_not": [
        "Не указывать один и тот же именованный аргумент более одного раза.",
        "Не использовать несуществующие имена параметров."
      ]
    },
    "examples": {
      "bad": "```vba\nSub CreateUser(login As String, isAdmin As Boolean)\nEnd Sub\n\nCreateUser login:=\"X\", Admin:=True     ' Неправильно: нет параметра Admin\nCreateUser login:=\"Y\", isAdmin:=True, isAdmin:=False   ' Неправильно: дублирование isAdmin\n```",
      "good": "```vba\nCreateUser login:=\"Alice\", isAdmin:=True   ' Правильно: точные имена без дублирования\n```"
    },
    "rationale": "VBA ожидает точного совпадения имени параметра. Ошибки в именовании приводят к синтаксической ошибке.",
    "tags": ["parameters", "named"]
  },
  {
    "uid": "parameters/mixed_named_positional",
    "title": "Смешивание именованных и позиционных аргументов",
    "severity": "compile_error",
    "vba_scope": ["StandardModule"],
    "context": "После именованного аргумента нельзя передавать позиционный – интерпретатор не сможет сопоставить значения.",
    "problem": "Позиционный аргумент следует только после неименованных, а не после именованных.",
    "cause": "После именованного аргумента передан позиционный.",
    "consequence": "Ошибка компиляции: синтаксическая ошибка в списке аргументов.",
    "rules": {
      "must": [
        "После начала указания именованных аргументов передавать все последующие аргументы тоже именованно."
      ],
      "must_not": [
        "Не использовать позиционные аргументы после именованных."
      ]
    },
    "examples": {
      "bad": "```vba\nSub FormatText(text As String, Optional bold As Boolean = False)\nEnd Sub\n\nFormatText \"Hello\", bold:=True, 255   ' Ошибка: после именованного идет безымянный\n```",
      "good": "```vba\nFormatText \"Hello\", bold:=True    ' Правильно: сначала позиционный \"Hello\", затем именованный bold\n```"
    },
    "rationale": "VBA не может сопоставить позиционный аргумент после именованного с конкретным параметром.",
    "tags": ["parameters", "named"]
  },
  {
    "uid": "validation/parameter_validation",
    "title": "Отсутствие проверки входных параметров",
    "severity": "warning",
    "vba_scope": ["StandardModule"],
    "context": "Параметры процедур/функций следует проверять на валидность заранее.",
    "problem": "Без проверки некорректный ввод (отрицательные числа, неверный диапазон) приводит к неверным расчётам.",
    "cause": "Не реализованы предусловия (preconditions) в начале процедуры.",
    "consequence": "Неверные результаты или сбои (например, отрицательная площадь).",
    "rules": {
      "must": [
        "В начале функции/процедуры проверять корректность входных значений (диапазон, положительность и т.д.).",
        "В случае неверных данных вызывать Err.Raise или выходить из процедуры."
      ],
      "must_not": [
        "Не выполнять вычисления без предварительной валидации."
      ]
    },
    "examples": {
      "bad": "```vba\nFunction RectangleArea(length As Double, width As Double) As Double\n    RectangleArea = length * width    ' Нет проверки: площадь не может быть отрицательной\nEnd Function\n```",
      "good": "```vba\nFunction RectangleArea(length As Double, width As Double) As Double\n    If length <= 0 Or width <= 0 Then\n        Err.Raise vbObjectError + 1, , \"Длина и ширина должны быть положительными\"\n    End If\n    RectangleArea = length * width\nEnd Function\n```"
    },
    "rationale": "Ранний выход при некорректных данных предотвращает распространение ошибок в коде (принцип fail-fast).",
    "tags": ["validation", "parameters"]
  },
  {
    "uid": "validation/null_empty",
    "title": "Отсутствие проверки Null/Empty",
    "severity": "runtime_error",
    "vba_scope": ["StandardModule"],
    "context": "Параметры могут быть Null (`Variant`) или пустыми строками (`String`).",
    "problem": "Конкатенация или операции с Null приводят к ошибке, а с пустой строкой – к логической ошибке.",
    "cause": "Не проверено, что переменная не Null или не пустая строка.",
    "consequence": "Runtime error (Invalid use of Null) или нежелательное поведение.",
    "rules": {
      "must": [
        "Использовать функции `IsNull`, `IsEmpty` или проверку на \"\" перед использованием таких параметров."
      ],
      "must_not": [
        "Не выполнять операции с переменными, которые могут быть Null, без проверки."
      ]
    },
    "examples": {
      "bad": "```vba\nSub GreetUser(userName As Variant)\n    MsgBox \"Здравствуйте, \" & userName   ' Ошибка, если userName = Null\nEnd Sub\n```",
      "good": "```vba\nSub GreetUser(userName As Variant)\n    If IsNull(userName) Or userName = \"\" Then\n        MsgBox \"Имя пользователя не задано.\"\n        Exit Sub\n    End If\n    MsgBox \"Здравствуйте, \" & userName\nEnd Sub\n```"
    },
    "rationale": "Проверка на Null/Empty предотвращает сбои при работе с базами данных и некорректными входными данными.",
    "tags": ["validation", "null"]
  },
  {
    "uid": "validation/collection_elements",
    "title": "Отсутствие проверки типов в коллекции",
    "severity": "runtime_error",
    "vba_scope": ["StandardModule"],
    "context": "Коллекция (`Collection`) VBA может содержать элементы разных типов.",
    "problem": "Операции над элементами без проверки типа могут привести к ошибке *Type Mismatch*.",
    "cause": "Смешанные типы в коллекции и использование их в арифметических/строковых операциях без проверки.",
    "consequence": "Runtime error (Type Mismatch).",
    "rules": {
      "must": [
        "Перед операциями с элементом коллекции проверять его тип (`IsNumeric`, `TypeName` и т.д.)."
      ],
      "must_not": [
        "Не предполагать, что все элементы коллекции одного типа без проверки."
      ]
    },
    "examples": {
      "bad": "```vba\nDim values As New Collection\nvalues.Add 100\nvalues.Add \"200\"\nDim total As Long, item As Variant\nFor Each item In values\n    total = total + item   ' Ошибка: сложение Long и String\nNext\n```",
      "good": "```vba\nDim total As Long, item As Variant\nFor Each item In values\n    If IsNumeric(item) Then\n        total = total + CLng(item)\n    Else\n        Debug.Print \"Пропущен нечисловой элемент: \" & TypeName(item)\n    End If\nNext\n```"
    },
    "rationale": "Проверка типов элементов защищает от ошибок и обеспечивает корректную обработку коллекции.",
    "tags": ["validation", "collections"]
  },
  {
    "uid": "validation/argument_type",
    "title": "Отсутствие проверки типа аргумента",
    "severity": "runtime_error",
    "vba_scope": ["StandardModule"],
    "context": "Если процедура ожидает определённый тип, передача другого без проверки опасна.",
    "problem": "Неявное преобразование или попытка операции с несоответствующим типом.",
    "cause": "Не указан ожидаемый тип параметра (Variant по умолчанию) или не проверен переданный объект.",
    "consequence": "Непредсказуемое поведение: неявные преобразования или ошибка при вызове метода.",
    "rules": {
      "must": [
        "Указывать ожидаемый тип параметра или проверять его внутри процедуры (TypeOf, IsObject и т.д.)."
      ],
      "must_not": [
        "Не полагаться на Variant/Object без проверки типа."
      ]
    },
    "examples": {
      "bad": "```vba\nSub PrintLength(val)\n    Debug.Print Len(val)    ' Если val не строка – неожиданный результат или ошибка\nEnd Sub\n\nPrintLength 123   ' Выведет 3 (число 123 преобразуется в \"123\"), возможно не то, что ожидалось\n```",
      "good": "```vba\nSub PrintLength(txt As String)\n    If txt = \"\" Then\n        Debug.Print \"Пустая строка\"\n    Else\n        Debug.Print Len(txt)\n    End If\nEnd Sub\n\nDim s As String\ns = CStr(123)\nPrintLength s  ' Явное приведение числа к строке перед передачей\n```"
    },
    "rationale": "Явные типы и проверки предотвращают скрытые ошибки и делают код более понятным.",
    "tags": ["validation", "types"]
  },
  {
    "uid": "optimization/enum_conversion",
    "title": "Лишняя конверсия Enum",
    "severity": "warning",
    "vba_scope": ["StandardModule"],
    "context": "Enum служат именованными константами и быстро сравниваются без конверсий.",
    "problem": "Частая конверсия Enum в другой тип (CInt, CStr) замедляет код и может вводить ошибки.",
    "cause": "Неверно предполагая, что нужно конвертировать Enum при сравнении.",
    "consequence": "Снижение производительности, возможный переполнение при конверсии.",
    "rules": {
      "must": [
        "Сравнивать Enum напрямую с константой перечисления."
      ],
      "must_not": [
        "Не использовать CInt, CStr и подобные для обращения к Enum без необходимости."
      ]
    },
    "examples": {
      "bad": "```vba\nEnum ColorMode\n    cmGray = 1\n    cmRGB = 2\nEnd Enum\n\nDim mode As ColorMode\nmode = cmGray\n\n' Плохой подход: лишнее преобразование Enum в число\nIf CInt(mode) = 2 Then\n    Debug.Print \"RGB\"\nEnd If\n```",
      "good": "```vba\nIf mode = cmRGB Then\n    Debug.Print \"RGB\"\nEnd If\n```"
    },
    "rationale": "Прямое использование Enum эффективнее и чище. Функции преобразования добавляют лишний оверхед.",
    "tags": ["performance", "enum"]
  },
  {
    "uid": "optimization/data_structure",
    "title": "Неправильный выбор массива/коллекции",
    "severity": "warning",
    "vba_scope": ["StandardModule"],
    "context": "Динамическое расширение массива в цикле (`ReDim Preserve`) очень неэффективно.",
    "problem": "Многократное `ReDim Preserve` или неиспользование `Collection` для динамического списка замедляет выполнение.",
    "cause": "Применение массивов без заранее известного размера или использование `Collection` при известном объёме.",
    "consequence": "Замедление алгоритма и избыточное использование памяти.",
    "rules": {
      "must": [
        "Если известен размер данных, сразу задавать массив этого размера.",
        "Если размер динамический, использовать `Collection`."
      ],
      "must_not": [
        "Не выполнять `ReDim Preserve` внутри больших циклов без крайней необходимости."
      ]
    },
    "examples": {
      "bad": "```vba\nDim arr() As Long\nFor i = 1 To 1000\n    ReDim Preserve arr(1 To i)\n    arr(i) = i\nNext i\n```",
      "good": "```vba\n' Использование массива с известным размером:\nDim arr(1 To 1000) As Long\nFor i = 1 To 1000\n    arr(i) = i\nNext i\n\n' Или использование коллекции для динамического списка:\nDim col As New Collection\nFor i = 1 To 1000\n    col.Add i\nNext i\n```"
    },
    "rationale": "Выделение памяти единовременно или использование коллекции устраняет многократные операции копирования данных.",
    "tags": ["performance", "arrays", "collections"]
  }
]

[
    {
        "uid": "M001",
        "title": "Неочистка ресурсов и утечки памяти",
        "severity": "High",
        "vba_scope": "Modules, Classes",
        "context": "Работа с объектами, коллекциями, массивами и внешними ресурсами в VBA.",
        "problem": "Отсутствие явной очистки и освобождения ресурсов после работы с данными. Например, большие массивы остаются в памяти даже после использования, объекты в коллекциях не освобождаются, ссылки на объекты не сбрасываются. В VBA сборка мусора происходит автоматически при выходе объектов из области видимости, но неправильное управление ссылками (особенно глобальными) может привести к утечкам. Некоторые объекты COM требуют явного закрытия методов (например, Recordset.Close для ADO или Workbook.Close для Excel). Игнорирование этих действий приводит к тому, что память и ресурсы (файлы, соединения) остаются занятыми.",
        "cause": "Неосвобождение объектов после использования: отсутствие `Set obj = Nothing` для очистки ссылок, отсутствие `Erase` для массивов, сохранение ссылок на объекты в глобальных переменных или в коллекциях, которые не очищаются. Игнорирование необходимости закрытия файлов и соединений.",
        "consequence": "Неосвобождённые ресурсы приводят к постепенному росту использования памяти и ресурсов приложением. Постепенно это может привести к ошибкам типа «Out of Memory» или ошибкам доступа, снижению производительности Excel и макросов, а также к «тёплому» зависанию приложения (теневым процессам). Открытые файлы или соединения могут оставаться заблокированными, что чревато сбоями в других частях системы.",
        "rules": {
            "must": [
                "Всегда освобождайте объекты после завершения их использования (`Set obj = Nothing`).",
                "Используйте `Erase` для очистки больших динамических массивов после работы.",
                "Закрывайте файлы и соединения сразу после работы (`Close` файловых дескрипторов, `Recordset.Close` и т.д.).",
                "Удаляйте объекты из коллекций или уничтожайте всю коллекцию (`Set collection = Nothing`), чтобы убрать все внутренние ссылки."
            ],
            "must_not": [
                "Не оставляйте объекты живыми в глобальных переменных без необходимости.",
                "Не полагайтесь только на выход процедур для освобождения объектов, особенно если ссылки хранятся в глобальном пространстве.",
                "Не игнорируйте необходимость вызова методов закрытия у COM-объектов (Workbook.Quit, Connection.Close и т.п.)."
            ]
        },
        "examples": {
            "bad": "Dim col As New Collection\nDim j As Long\nFor j = 1 To 1000\n    Dim rng As Range\n    Set rng = Worksheets(1).Cells(j, 1)\n    col.Add rng  ' сохраняем ссылки на Range в коллекции\nNext j\n' ... забыли очистить col или не используем её далее",
            "good": "Dim fnum As Integer\nfnum = FreeFile\nOpen \"data.txt\" For Output As #fnum\n' ... запись данных ...\nClose #fnum  ' файл закрыт\n\nDim dict As Object\nSet dict = CreateObject(\"Scripting.Dictionary\")\n' ... заполнение словаря ...\nSet dict = Nothing  ' словарь и его содержимое освобождены\n\nDim bigArr() As Double\nReDim bigArr(1 To 1000000)\n' ... использование bigArr ...\nErase bigArr  ' освобождение памяти под массив"
        },
        "rationale": "В VBA объекты управляются COM-счётчиком ссылок: пока есть хотя бы одна ссылка на объект, он остаётся в памяти. Если не сбрасывать ссылки вручную, объекты и связанные ресурсы (файлы, соединения) будут удерживаться до выгрузки проекта. Это приводит к утечкам памяти и ресурсов, снижает надёжность и производительность приложения.",
        "tags": ["memory", "resource-leak", "performance", "COM", "cleanup"]
    },
    {
        "uid": "M002",
        "title": "Неправильное управление жизненным циклом объектов",
        "severity": "High",
        "vba_scope": "Classes, Modules",
        "context": "Создание и использование объектов (особенно классов) в VBA.",
        "problem": "Объекты создаются, используются и остаются в памяти дольше, чем нужно, или уничтожаются раньше времени. Примеры: сохранение ссылок на объект в глобальных переменных, образование циклических ссылок между объектами (когда объект A содержит ссылку на B, а B — на A), неявное удаление объекта, который всё ещё нужен.",
        "cause": "Избыточное использование глобальных объектов без освобождения, двусторонние ссылки между объектами без явного разрыва, несоответствие области видимости переменных и ожидаемого времени жизни объектов.",
        "consequence": "Неосвобождённые объекты удерживают ресурсы и приводят к утечкам памяти. Циклические ссылки не разрушаются сборщиком VBA, поэтому объекты никогда не удаляются, даже после выхода из процедур. С другой стороны, преждевременное удаление объекта может вызвать ошибку «Object variable not set» при попытке к нему обратиться.",
        "rules": {
            "must": [
                "Избегайте циклических ссылок между объектами. Если надо связать объекты двусторонне, предусмотрите механизм разрыва ссылки.",
                "При необходимости разрывать циклы, явно устанавливайте `obj = Nothing` для одной или обеих сторон.",
                "Ограничивайте область видимости объектов — создавайте их там, где используются, и освобождайте сразу после окончания работы.",
                "Если нужно длительное существование объекта (например, глобальное соединение), документируйте и централизуйте его создание и уничтожение (например, в `Initialize`/`Cleanup`)."
            ],
            "must_not": [
                "Не держите лишние ссылки на объект при выходе из процедуры (например, в локальной переменной класса), полагаясь на автоматическое уничтожение.",
                "Не используйте глобальные объекты без необходимости — они живут до закрытия файла или проекта.",
                "Не полагайтесь на `Class_Terminate` для разрыва циклических ссылок — он не сработает, если цикл существует."
            ]
        },
        "examples": {
            "bad": "' Класс Parent\nPublic Child As Child\n\n' Класс Child\nPublic Parent As Parent\n\n' --- где-то в коде ---\nDim p As New Parent\nDim c As New Child\nSet p.Child = c\nSet c.Parent = p\n' Конец процедуры – p и c выходят из области видимости\n",
            "good": "Dim xlApp As Object\nSet xlApp = CreateObject(\"Excel.Application\")\nxlApp.Workbooks.Open \"report.xlsx\"\n' ... работа с Excel ...\nxlApp.Quit\nSet xlApp = Nothing  ' Явно закрываем и удаляем приложение Excel"
        },
        "rationale": "VBA использует COM-счётчик ссылок для управления временем жизни объектов. Если на объект никто не ссылается, он автоматически уничтожается. Глобальные объекты или циклические ссылки нарушают это правило, вследствие чего объекты либо не удаляются, либо удаляются когда это не ожидается. Правильное управление жизненным циклом объектов (создание, область видимости, уничтожение) позволяет избежать утечек и нестабильного поведения.",
        "tags": ["lifecycle", "memory", "object", "COM", "cleanup"]
    },
    {
        "uid": "CS001",
        "title": "Чрезмерно глубокая цепочка вызовов процедур",
        "severity": "Medium",
        "vba_scope": "Modules, Classes",
        "context": "Архитектура кода с множеством уровней вызовов процедур.",
        "problem": "Для выполнения задачи управление проходит через слишком много уровней вложенных вызовов процедур. Например, процедура A вызывает B, та вызывает C и т.д., создавая длинную цепочку. Это может быть результатом избыточной декомпозиции, ненужных обёрток или многослойной абстракции.",
        "cause": "Непродуманная декомпозиция, чрезмерное использование шаблонов или абстракций, перекладывание работы от функции к функции без необходимости.",
        "consequence": "Слишком глубокая цепочка затрудняет понимание и отладку кода: программисту приходится переходить через множество функций-посредников, чтобы найти реальную логику. Поддержка усложняется, поскольку нужно учитывать много уровней вызовов. Это также снижает производительность: каждый вызов процедуры добавляет накладные расходы (особенно в VBA без inline-оптимизации). В экстремальных случаях возможен переполнение стека.",
        "rules": {
            "must": [
                "Минимизируйте глубину цепочек вызовов: каждый уровень должен добавлять обоснованную ценность.",
                "При проектировании алгоритма старайтесь выполнять задачи как можно ближе к точке, где они нужны, избегая лишних посредников.",
                "Используйте явное именование процедур, отражающее их ответственность, чтобы не вводить ненужные слои абстракции."
            ],
            "must_not": [
                "Не создавайте лишние обёртки над процедурами без веской причины.",
                "Не перекладывайте работу на новую процедуру, если это не упрощает логику; слишком короткие процедуры для одной операции могут усложнить структуру.",
                "Не допускайте рекурсии без чётких условий выхода."
            ]
        },
        "examples": {
            "bad": "Sub Level1()\n    Level2\nEnd Sub\n\nSub Level2()\n    Level3\nEnd Sub\n\nSub Level3()\n    Level4\nEnd Sub\n\n' ... и так далее ...\n\nSub Level9()\n    Level10\nEnd Sub\n\nSub Level10()\n    ' Наконец-то выполняем нужную работу\n    Debug.Print \"Done!\"\nEnd Sub",
            "good": "Sub DoActualWork()\n    ' Логика выполняется здесь без ненужных посредников\n    Debug.Print \"Done!\"\nEnd Sub"
        },
        "rationale": "Глубокие цепочки вызовов усложняют архитектуру и снижают читаемость кода. Программисты теряют время на понимание многочисленных функций-посредников, и это увеличивает вероятность ошибок при изменениях. Каждая дополнительная процедура добавляет накладные расходы при выполнении.",
        "tags": ["architecture", "complexity", "performance", "refactoring"]
    },
    {
        "uid": "CS002",
        "title": "Дублирование кода в классах из-за неправильной организации интерфейсов",
        "severity": "Medium",
        "vba_scope": "Classes",
        "context": "Определение методов в классах, реализующих общие интерфейсы без наследования.",
        "problem": "В VBA нет наследования реализации, поэтому для общего функционала в нескольких классах часто дублируют код. Например, два класса, реализующие один интерфейс, копируют одинаковую логику в своих обработчиках. Общий код не вынесен в единое место, а повторяется в каждом классе.",
        "cause": "Ограничение VBA в отсутствии наследования, приводящее к копированию/вставке кода. Отсутствие единых утилит или модулей для общих операций, а также неправильное проектирование интерфейсов.",
        "consequence": "Любое изменение общей логики требует правки в нескольких местах, что чревато несинхронными изменениями и ошибками. Код становится труднопросматриваемым и тяжёлым в сопровождении. Ошибки в одной реализации могут не быть исправлены в остальных.",
        "rules": {
            "must": [
                "Выносите общий код в отдельные процедуры/модули или используйте композицию для повторного использования.",
                "Используйте общие функции или менеджеры для поведения, общего для нескольких классов.",
                "Проектируйте интерфейсы так, чтобы делегировать общую логику (например, паттерн Стратегии или Утилит)."
            ],
            "must_not": [
                "Не копируйте один и тот же блок кода в разных классах.",
                "Не добавляйте новые методы в интерфейс, если можно обеспечить общую реализацию извне.",
                "Не оставляйте `NotImplemented` (пустые) методы: это указывает на дублирование."
            ]
        },
        "examples": {
            "bad": "' Интерфейс IAnimal с методом Eat\nPublic Sub Eat(Food As String)\nEnd Sub\n\n' Класс Cat реализует IAnimal\nImplements IAnimal\nPrivate Sub IAnimal_Eat(Food As String)\n    If Food = \"fish\" Then\n        Debug.Print \"Cat eats the fish.\"\n    Else\n        Debug.Print \"Cat refuses the food.\"\n    End If\nEnd Sub\n\n' Класс Dog реализует IAnimal\nImplements IAnimal\nPrivate Sub IAnimal_Eat(Food As String)\n    If Food = \"meat\" Then\n        Debug.Print \"Dog eats the meat.\"\n    Else\n        Debug.Print \"Dog refuses the food.\"\n    End If\nEnd Sub",
            "good": "' Общая процедура в модуле для выполнения Eat\nPublic Sub FeedAnimal(AnimalType As String, Food As String)\n    If AnimalType = \"Cat\" And Food = \"fish\" Then\n        Debug.Print \"Cat eats the fish.\"\n    ElseIf AnimalType = \"Dog\" And Food = \"meat\" Then\n        Debug.Print \"Dog eats the meat.\"\n    Else\n        Debug.Print \"Animal refuses the food.\"\n    End If\nEnd Sub"
        },
        "rationale": "Поскольку VBA-классы не поддерживают наследование, дублирование кода — частая ошибка. Вынося общие фрагменты в отдельные модули или процедуры, легко гарантировать консистентность поведения. Общая реализация упрощает поддержку и уменьшает технический долг.",
        "tags": ["design", "code-duplication", "classes", "architecture"]
    },
    {
        "uid": "INIT001",
        "title": "Неправильная попытка использовать класс как точку входа",
        "severity": "Medium",
        "vba_scope": "Classes, Modules",
        "context": "Определение начальной точки запуска приложения или модуля.",
        "problem": "Размещение основного кода инициализации или главной процедуры внутри класса (Class Module) вместо обычного модуля. Например, попытка использовать `Class_Initialize` или создать `Sub Start` в классе как точку входа макроса.",
        "cause": "Непонимание модели VBA: код в классовом модуле не выполняется автоматически без создания экземпляра класса. Разработчики пытаются поместить точку входа в класс, но не создают объект класса.",
        "consequence": "Код в классе не будет выполнен, пока явно не будет создан экземпляр этого класса. Макрос может ничего не делать или не запускаться вовсе без явного вызова, что приводит к скрытым ошибкам и непредсказуемому поведению.",
        "rules": {
            "must": [
                "Размещайте стартовую логику (точку входа) в стандартном модуле, а не в классе.",
                "Создавайте экземпляры классов только когда это действительно необходимо, а не для имитации глобальных точек входа.",
                "Используйте процедуры `Sub` или `Function` в обычных модулях как видимые точки входа макросов."
            ],
            "must_not": [
                "Не пытайтесь вызывать процедуры из класса напрямую как макросы без создания объекта.",
                "Не используйте `Class_Initialize` как способ автоматически запустить основной код приложения.",
                "Не объявляйте `Public Sub` внутри класса без создания его экземпляра при попытке запустить."
            ]
        },
        "examples": {
            "bad": "' Модуль класса AppManager (неправильно – попытка использовать класс как точку входа)\nOption Explicit\nPrivate Sub Class_Initialize()\n    InitializeApp  ' Попытка вызвать главный процесс при создании объекта\nEnd Sub\n\nPublic Sub InitializeApp()\n    ' Код инициализации приложения\n    Debug.Print \"Initializing application.\"\nEnd Sub\n\nPublic Sub RunMainTask()\n    Debug.Print \"Running main task.\"\nEnd Sub",
            "good": "' Стандартный модуль: функция Main как точка входа\nPublic Sub Main()\n    Dim mgr As New AppManager\n    mgr.InitializeApp\n    mgr.RunMainTask\nEnd Sub\n\n' Класс AppManager без точки входа\nOption Explicit\nPublic Sub InitializeApp()\n    Debug.Print \"Initializing application.\"\nEnd Sub\nPublic Sub RunMainTask()\n    Debug.Print \"Running main task.\"\nEnd Sub"
        },
        "rationale": "VBA не запускает код из классов без явного создания объекта. Стандартные модули доступны глобально, поэтому их используют для точек входа и утилит. Размещение запускающего кода в классе без экземпляра приведет к тому, что основная логика никогда не выполнится.",
        "tags": ["initialization", "entry-point", "classes", "design"]
    },
    {
        "uid": "PERF001",
        "title": "Отсутствие буферизации и кэширования при больших объёмах данных",
        "severity": "High",
        "vba_scope": "Modules",
        "context": "Обработка больших диапазонов данных и взаимодействие с внешними источниками (листы, файлы, базы).",
        "problem": "Код обрабатывает большие наборы данных без промежуточного сохранения результатов, обращаясь напрямую к медленным источникам (листам Excel, файлам, базам данных) на каждой итерации. Например, чтение/запись ячеек внутри большого цикла вместо единовременных операций.",
        "cause": "Непонимание стоимости операций ввода/вывода: чтение из листа или диска в цикле происходит многократно без использования памяти. Отсутствие временных массивов или переменных для часто повторяющихся вычислений.",
        "consequence": "Производительность серьёзно падает: код с простыми операциями может работать в десятки раз дольше. Операции с Excel или файлами в цикле приводят к множеству накладных вызовов, замедляя выполнение. Повторное вычисление функций с теми же аргументами увеличивает нагрузку на процессор и может приводить к устаревшим данным.",
        "rules": {
            "must": [
                "Считывайте данные блоками (например, диапазон в массив) и обрабатывайте в памяти вместо поклеточных обращений.",
                "Буферизуйте результаты дорогих операций (запоминайте их в переменные или коллекции) при повторном использовании.",
                "Используйте массивы или словари (`Scripting.Dictionary`) для накопления и обработки данных вместо множественных обращений к рабочему листу."
            ],
            "must_not": [
                "Не обрабатывайте данные одну запись за другой на листе без необходимости.",
                "Не вызывайте одни и те же функции с одинаковыми параметрами в циклах, не сохраняя их результат.",
                "Не записывайте результат по одной ячейке за раз, если можно записать диапазон одномоментно."
            ]
        },
        "examples": {
            "bad": "Dim total As Double\nDim i As Long\ntotal = 0\nFor i = 1 To 1000\n    total = total + Cells(i, 1).Value  ' Чтение с листа каждую итерацию\nNext i\nMsgBox \"Сумма = \" & total",
            "good": "Dim total As Double\nDim data As Variant\nDim i As Long\ndata = Range(Cells(1, 1), Cells(1000, 1)).Value  ' Чтение 1000 ячеек за один вызов\ntotal = 0\nFor i = 1 To 1000\n    total = total + data(i, 1)\nNext i\nMsgBox \"Сумма = \" & total"
        },
        "rationale": "Операции ввода/вывода, такие как обращение к Excel, файлам или базам, являются относительно медленными. Буферизация данных (чтение/запись блоками) позволяет существенно сократить количество таких операций и ускорить выполнение.",
        "tags": ["performance", "buffering", "caching", "optimization", "excel"]
    },
    {
        "uid": "LOG001",
        "title": "Вложенные логические конструкции вместо циклов",
        "severity": "Medium",
        "vba_scope": "Modules",
        "context": "Условные конструкции, которые повторяют одинаковую логику вместо использования цикла.",
        "problem": "Код дублирует одну и ту же логику путём ручного повторения блоков условий (If/Else) для каждого элемента, вместо того чтобы использовать цикл. Например, несколько последовательных `If a(i) > 0 Then` вместо `For`.",
        "cause": "Непонимание или нежелание использовать циклы, использование 'копипаста' при добавлении новых проверок. Ограниченный опыт разработчика приводит к дублированию условий для каждого индекса.",
        "consequence": "Код становится громоздким и трудночитаемым. Любое изменение алгоритма требует исправления во множестве мест. Увеличивается вероятность ошибок (опечатки в индексах) и снижается производительность при большом количестве проверок. Поддержка такого кода затруднена.",
        "rules": {
            "must": [
                "Используйте циклы (`For`, `For Each`) для повторяющихся действий над множеством элементов.",
                "Если код в нескольких `If` отличается только индексом или именем переменной, замените дублирование на цикл.",
                "Выносите повторяющийся код в отдельную функцию или подпрограмму, чтобы избежать копипаста."
            ],
            "must_not": [
                "Не дублируйте однотипные условия или действия вручную для каждого элемента.",
                "Не используйте глубокую вложенность `If` без необходимости для последовательной проверки схожих условий.",
                "Не оставляйте код, отличающийся только номером индекса — используйте цикл с переменной."
            ]
        },
        "examples": {
            "bad": "' Есть массив a(1..3), нужно вывести положительные элементы\nIf a(1) > 0 Then Debug.Print a(1)\nIf a(2) > 0 Then Debug.Print a(2)\nIf a(3) > 0 Then Debug.Print a(3)",
            "good": "Dim i As Integer\nFor i = LBound(a) To UBound(a)\n    If a(i) > 0 Then Debug.Print a(i)\nNext i"
        },
        "rationale": "Принцип DRY (Don't Repeat Yourself) нарушается, когда одинаковая логика копируется вручную. Цикл позволяет обрабатывать произвольное количество элементов компактно и гибко. Использование цикла облегчает сопровождение и улучшает масштабируемость кода.",
        "tags": ["logic", "refactoring", "best-practice", "readability"]
    },
    {
        "uid": "ERR001",
        "title": "Неконтролируемое использование `Exit` и `On Error Resume Next`",
        "severity": "High",
        "vba_scope": "Modules",
        "context": "Обработка ошибок и управление потоком в процедурах.",
        "problem": "Преждевременный выход из процедур (`Exit Sub`, `Exit Function`) до выполнения всех необходимых действий (например, очистки ресурсов), а также бесконтрольное подавление ошибок (`On Error Resume Next`) без их обработки. Часто `Resume Next` включается и не отключается.",
        "cause": "Несоблюдение структуры блоков обработки: ранний `Exit` пропускает код последующей очистки, а `On Error Resume Next` используется вместо избирательной обработки ошибок.",
        "consequence": "При `Exit` ресурсы (файлы, соединения) могут остаться не закрытыми, что вызывает утечки. `On Error Resume Next` скрывает реальные ошибки: код продолжает работу в некорректном состоянии, это затрудняет отладку и может привести к скрытым багам и сбоям в других местах.",
        "rules": {
            "must": [
                "При выходе из процедуры убедитесь, что все ресурсы освобождены (например, в блоке `Cleanup`).",
                "Используйте `On Error ... GoTo` для локальной обработки ошибок и всегда отключайте `Resume Next` после нужной операции.",
                "Если необходимо игнорировать ожидаемые ошибки, сразу после них делайте `On Error GoTo 0` или аналогичную обработку."
            ],
            "must_not": [
                "Не используйте `Exit Sub/Function` или `Resume Next` без необходимости чистки ресурсов перед ними.",
                "Не оставляйте `Resume Next` включённым на весь остаток процедуры.",
                "Не полагайтесь на `On Error Resume Next` как на обычный механизм обработки — он скрывает реальные проблемы."
            ]
        },
        "examples": {
            "bad": "Sub WriteData(ByVal text As String, ByVal fileName As String)\n    Dim fnum As Integer\n    fnum = FreeFile\n    Open fileName For Output As #fnum\n    If text = \"\" Then\n        Exit Sub  ' выходим, не закрыв файл\n    End If\n    Print #fnum, text\n    Close #fnum\nEnd Sub",
            "good": "Sub WriteData(ByVal text As String, ByVal fileName As String)\n    Dim fnum As Integer\n    fnum = FreeFile\n    Open fileName For Output As #fnum\n    If text <> \"\" Then\n        Print #fnum, text\n    End If\n    Close #fnum\nEnd Sub"
        },
        "rationale": "`Exit` досрочно завершает процедуру, пропуская дальнейшую логику (например, очистку), а `On Error Resume Next` полностью подавляет ошибки. Оба подхода приводят к ошибочному поведению: ресурсы могут остаться открытыми, а скрытые ошибки в дальнейшем приведут к отказам или некорректной работе. Правильный подход — контролировать точки выхода и ошибки явно.",
        "tags": ["error-handling", "code-quality", "resources", "control-flow"]
    },
    {
        "uid": "EVT001",
        "title": "Некорректное управление событиями и наблюдателями",
        "severity": "High",
        "vba_scope": "Modules, Class Modules",
        "context": "Работа с событиями Excel и объектными наблюдателями (`WithEvents`).",
        "problem": "Неправильное включение/выключение обработки событий и управление подписчиками (обработчиками). Типичные ошибки: неотключение событий при массовых операциях, забывчивое восстановление `Application.EnableEvents`, создание нескольких обработчиков на одно событие или неосвобождение объектов-субъектов или `WithEvents` объектов.",
        "cause": "Игнорирование влияния событий на производительность и логику, множественные регистрации обработчиков без снятия, отсутствие контроля над глобальным состоянием флага событий.",
        "consequence": "Макросы могут резко замедлиться из-за массовых вызовов обработчиков. Забытые отключенные события приводят к тому, что последующие действия пользователя не отслеживаются. Множественные слушатели могут выполнять код несколько раз или приводить к утечкам (объекты не освобождаются). В худшем случае возможны бесконечные циклы событий и зависание Excel.",
        "rules": {
            "must": [
                "Отключайте события (`Application.EnableEvents = False`) перед массовыми изменениями и всегда включайте обратно.",
                "В обработчиках событий избегайте изменений, которые снова триггерят то же событие (в случае необходимости, отключайте события на время).",
                "Регистрация обработчиков (`WithEvents`) должна производиться один раз; не создавайте новый слушатель при каждом вызове.",
                "При завершении работы снимайте обработчики (`Set evt = Nothing` или `EnableEvents = True`) во `Workbook_BeforeClose` или аналогичных местах."
            ],
            "must_not": [
                "Не оставляйте `EnableEvents = False` включённым после завершения макроса.",
                "Не создавайте дублирующие экземпляры класса-слушателя без снятия предыдущих.",
                "Не отключайте события без явного последующего включения — это приведёт к «онемению» интерфейса."
            ]
        },
        "examples": {
            "bad": "Private Sub Worksheet_Change(ByVal Target As Range)\n    ' Каждый раз при изменении любой ячейки копируем значение в соседнюю колонку\n    If Not Intersect(Target, Range(\"A:A\")) Is Nothing Then\n        Target.Offset(0, 1).Value = Target.Value * 2\n    End If\nEnd Sub\n\nSub FillData()\n    Dim i As Long\n    For i = 1 To 1000\n        Cells(i, 1).Value = i  ' это триггерит Worksheet_Change 1000 раз!\n    Next i\nEnd Sub",
            "good": "Sub FillDataOptimized()\n    Application.EnableEvents = False  ' временно отключаем события\n    Dim i As Long\n    For i = 1 To 1000\n        Cells(i, 1).Value = i\n    Next i\n    Application.EnableEvents = True   ' включаем события обратно\nEnd Sub"
        },
        "rationale": "События в Excel могут вызываться часто и непредсказуемо. Неконтролируемое управление событиями ведёт к замедлению и ошибкам. Правильное использование `EnableEvents` и контроля слушателей позволяет избежать избыточных вызовов, дублирования операций и скрытых сбоев.",
        "tags": ["events", "performance", "Excel", "best-practice", "reliability"]
    }
]

````
