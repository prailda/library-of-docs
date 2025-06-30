Отлично, я подготовлю пошаговое руководство по созданию собственного аналога функции Deep Research от OpenAI с использованием моделей o3 (OpenAI), Claude, DeepSeek, а также Python и VS Code. Я подберу оптимальный стек технологий — например, LangChain или простую Python-оркестрацию — для запуска в консольном режиме, без необходимости создания пользовательского интерфейса. В инструкции будет подробное описание всех шагов, включая установку зависимостей, структуру проекта, реализацию механизма оркестрации, подключение моделей через API и запуск исследования.

Я сообщу, когда всё будет готово.


# Аналог Deep Research на Python с OpenAI, Claude и DeepSeek

Ниже приведён подробный пошаговый гайд по созданию **консольного приложения**, имитирующего функцию “Deep Research” от OpenAI. Мы будем использовать модели OpenAI (версии o3), Anthropic Claude и DeepSeek, разбивая задачу на итеративные этапы, управляемые *оркестратором*. Рекомендуемые инструменты: **LangChain** (для оркестрации LLM), **Pydantic** (для настройки) и пр. Код будет приводиться с комментариями, чтобы его можно было копировать и запускать.

## 1. Подготовка окружения и зависимостей

* **Установите Python 3.9+** (рекомендуется 3.10+). Создайте виртуальное окружение (venv).

* Установите необходимые библиотеки через `pip`:

  ```bash
  pip install openai anthropic requests python-dotenv pydantic langchain
  ```

  * `openai` – для OpenAI API (GPT-модели).
  * `anthropic` – для работы с Claude (Anthropic).
  * `requests` – понадобится для вызова DeepSeek API.
  * `python-dotenv` – для загрузки переменных окружения из `.env`.
  * `pydantic` – для упрощённой настройки (валидирует конфигурацию).
  * `langchain` (опционально) – фреймворк для оркестрации LLM; упрощает создание цепочек и агентов.

* **Создайте файл `.env`** в корне проекта и добавьте туда ключи доступа к API (получите их в личных кабинетах OpenAI, Anthropic и DeepSeek):

  ```
  OPENAI_API_KEY=sk-...
  ANTHROPIC_API_KEY=sk-...
  DEEPSEEK_API_KEY=sk-...
  ```

  Этот файл прочитаем в конфигурации с помощью Pydantic/BaseSettings.

* По аналогии с [Anthropic-документацией](https://docs.anthropic.com) можно настроить `ANTHROPIC_API_KEY` в окружении, а [DeepSeek API](https://api-docs.deepseek.com/) совместим с форматом OpenAI-API. Таким образом мы получим доступ ко всем трем моделям.

## 2. Настройка API-доступа

Используем `pydantic` для централизованной загрузки настроек:

```python
# config.py
from pydantic import BaseSettings

class Settings(BaseSettings):
    openai_api_key: str
    anthropic_api_key: str
    deepseek_api_key: str

    class Config:
        env_file = ".env"  # автоматически загрузит переменные окружения из .env

settings = Settings()
```

Теперь `settings.openai_api_key`, `settings.anthropic_api_key` и `settings.deepseek_api_key` содержат соответствующие ключи. В коде мы установим их в клиентов API:

```python
import os
import openai
import anthropic

# Загрузка ключей из настроек
import config
os.environ["OPENAI_API_KEY"] = config.settings.openai_api_key
os.environ["ANTHROPIC_API_KEY"] = config.settings.anthropic_api_key

# Инициализация клиентов OpenAI и Anthropic
openai.api_key = os.getenv("OPENAI_API_KEY")
client_claude = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
```

Для DeepSeek сделаем простой HTTP-запрос через `requests`. Мы используем их совместимый OpenAI-формат API (base\_url `https://api.deepseek.com/v1`):

```python
import requests

DEESEEK_URL = "https://api.deepseek.com/v1/chat/completions"
DEESEEK_HEADERS = {
    "Content-Type": "application/json",
    "Authorization": f"Bearer {config.settings.deepseek_api_key}"
}
```

## 3. Архитектура проекта (оркестратор + воркеры)

Проект предлагается разделить на несколько модулей:

```
deep_research/
├── config.py              # Конфигурация и загрузка ключей (Pydantic)
├── orchestrator.py        # Основной код-оркестратор
└── workers/
    ├── openai_worker.py   # Вызов OpenAI (GPT-o3)
    ├── anthropic_worker.py# Вызов Claude (Anthropic)
    └── deepseek_worker.py # Вызов DeepSeek API
```

* **Оркестратор (`orchestrator.py`)** – управляет пошаговой логикой, вызывает воркеры и обрабатывает промежуточные результаты. Он получает на вход запрос пользователя, планирует этапы и сводит финальный отчёт.
* **Воркеры** – модули, реализующие взаимодействие с конкретной моделью или сервисом. Каждый воркер экспортирует функцию, принимающую текст-промпт и возвращающую ответ. Например, `ask_openai(prompt)`, `ask_claude(prompt)`, `ask_deepseek(prompt)`.

Такая структура позволяет легко масштабировать систему и, при желании, вставить вместо консоли веб-интерфейс на FastAPI или любой другой инструмент, разделив логику по сервисам. Для создания агентных цепочек можно применять LangChain – здесь мы реализуем простую версию вручную.

## 4. Реализация кода

### 4.1. Настройка воркеров

**`workers/openai_worker.py`** – вызывает OpenAI-модель GPT (версия o3). Для примера используем `gpt-4o3` (в реальности может быть `gpt-4o-mini`, если доступна) и чат-формат:

```python
# workers/openai_worker.py
import openai

# Предполагается, что ключ OpenAI уже установлен в openai.api_key
openai.api_base = "https://api.openai.com/v1"

def ask_openai(prompt: str) -> str:
    """
    Отправляет запрос в OpenAI ChatCompletion (модель o3).
    Возвращает сгенерированный текст.
    """
    response = openai.ChatCompletion.create(
        model="gpt-4o3",  # модель OpenAI версии o3
        messages=[{"role": "user", "content": prompt}],
        temperature=0.7
    )
    # Читаем ответ из первого выбора
    return response.choices[0].message.content
```

**`workers/anthropic_worker.py`** – вызывает Claude с помощью официальной библиотеки Anthropic:

```python
# workers/anthropic_worker.py
import anthropic

# Клиент использует ключ из переменной ANTHROPIC_API_KEY по умолчанию
client = anthropic.Anthropic(
    api_key=anthropic.os.getenv("ANTHROPIC_API_KEY")
)

def ask_claude(prompt: str) -> str:
    """
    Отправляет запрос в Claude (Anthropic) и возвращает ответ.
    """
    response = client.messages.create(
        model="claude-3-7b-sonnet-20250219",  # или другая доступная версия Claude 3 Opus
        max_tokens=1000,
        messages=[{"role": "user", "content": prompt}]
    )
    # Метод возвращает объект с полем .content
    return response.content
```

**`workers/deepseek_worker.py`** – делает HTTP-запрос к DeepSeek API:

```python
# workers/deepseek_worker.py
import requests
from config import settings

DEESEEK_URL = "https://api.deepseek.com/v1/chat/completions"

def ask_deepseek(prompt: str) -> str:
    """
    Отправляет запрос в DeepSeek API и возвращает ответ.
    DeepSeek совместим по API с OpenAI:contentReference[oaicite:6]{index=6}.
    """
    headers = {
        "Authorization": f"Bearer {settings.deepseek_api_key}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": "deepseek-chat",  # DeepSeek-V3
        "messages": [{"role": "user", "content": prompt}]
    }
    res = requests.post(DEESEEK_URL, json=payload, headers=headers)
    res.raise_for_status()
    # Получаем текст из первого выбора ответа
    return res.json()["choices"][0]["message"]["content"]
```

### 4.2. Оркестратор и логика шагов

Файл **`orchestrator.py`** объединяет всё вместе. Здесь мы показываем пример упрощённого *итеративного механизма*. На практике Deep Research строит сложный план; мы же вручную зададим несколько шагов:

```python
# orchestrator.py
from config import settings
from workers.openai_worker import ask_openai
from workers.anthropic_worker import ask_claude
from workers.deepseek_worker import ask_deepseek

def main():
    # Ввод вопроса пользователя
    question = input("Введите вопрос для исследования: ").strip()
    if not question:
        print("Ошибка: запрос не введён.")
        return

    print("\n[1/3] Формируем план исследования...")
    plan_prompt = (
        f"Задача: {question}\n"
        "Опишите план многошагового веб-исследования, разделённый на пункты."
    )
    plan = ask_openai(plan_prompt)
    print(f"План исследования:\n{plan}\n")

    # Шаг 1: получение дополнительной информации (симулированный поиск)
    print("[2/3] Выполняем поиск и сбор информации...")
    # Используем DeepSeek как "поисковый" шага
    search_prompt = f"Выполните поиск по плану: {plan}\nСоберите ключевую информацию."
    info = ask_deepseek(search_prompt)
    print(f"Информация от DeepSeek:\n{info}\n")

    # Шаг 2: анализ и уточнение (Claude может свести, дополнить)
    print("[3/3] Анализируем и формируем выводы...")
    analyze_prompt = (
        f"Основываясь на планировании и найденной информации:\n{plan}\n{info}\n"
        "Дайте конечный краткий отчёт."
    )
    summary = ask_claude(analyze_prompt)
    print(f"Окончательный отчёт от Claude:\n{summary}\n")

if __name__ == "__main__":
    main()
```

**Комментарии к коду**:

* Оркестратор выполняет три этапа: планирование, сбор информации и финальный анализ. Это упрощённый пример; вы можете добавить дополнительные шаги и корректировки.
* На первом шаге мы просим OpenAI сформировать план действий по вопросу. Здесь моделируется понимание задачи.
* Второй шаг использует DeepSeek для получения сведений по плану (в реальном решении здесь мог бы быть веб-скрейпинг или расширенный поиск). DeepSeek API совместим с OpenAI-форматом, поэтому мы используем его как “модель поиска”.
* Третий шаг – обращение к Claude, который на основе плана и найденной информации выдаёт итоговый отчёт. Параллельно можно вставить дополнительные шаги с другими моделями, например, проверки ответа GPT или добавления деталей.
* Весь ввод/вывод осуществляется через консоль (`input`/`print`), без веб-интерфейса.

## 5. Запуск, тестирование и отладка

* **Запуск приложения**: Убедитесь, что установлены все зависимости и настроены `.env`. В корне проекта выполните:

  ```bash
  python orchestrator.py
  ```

  Введите в консоли вопрос для исследования. Пример:

  ```
  Введите вопрос для исследования: Лучшая игровая клавиатура до 100$ в 2025 году
  ```

  Программа пошагово выведет план, собранную информацию и итоговый отчёт.

* **Тестирование**: Попробуйте разные запросы и следите за промежуточными выводами. Это поможет понять, как модель формирует план и какие ответы генерируются. Для отладки можно увеличить количество токенов или менять `temperature`.

* **Отладка**:

  * Если возникает ошибка API, убедитесь, что ключи верны и не исчерпан лимит.
  * Для детального логирования можно печатать ответы `response` целиком или сохранять их в файл.
  * Помните, что Claude и DeepSeek работают по-разному: у Claude потоковая генерация, у DeepSeek – просто статика.

## 6. Заключение

В результате у вас должно получиться консольное приложение, где на основе заданного вопроса **оркестратор** запускает последовательность запросов к разным моделям (OpenAI o3, Claude, DeepSeek) и формирует полный отчёт. Такой многоступенчатый подход позволяет глубже исследовать тему (аналогично OpenAI Deep Research), комбинируя сильные стороны разных LLM. При необходимости этот прототип можно расширить с помощью LangChain и LangGraph для более сложной логики, FastAPI – для API или Pydantic – для строгой модели данных.

**Полезные ссылки**: [Deep Research (OpenAI)](https://openai.com/index/introducing-deep-research), [Anthropic Claude API](https://docs.anthropic.com), [DeepSeek API Docs](https://api-docs.deepseek.com), [LangChain Docs](https://python.langchain.com).
