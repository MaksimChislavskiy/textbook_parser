# Парсер сборника задач по математике (5-6 классы)
Парсер для извлечения задач, ответов и оглавления из документа Microsoft Word (`tekstovye_zadachi_po_matematike.docx`) и объединения их в структурированный Excel-файл.
## Установка

1. Клонируй репозиторий:
   ```bash
   git clone git@github.com:MaksimChislavskiy/textbook_parser.git
   cd textbook_parser
2. Создай виртуальное окружение и активируй:
    ```bash
    python -m venv venv
    source venv/Scripts/activate  # Windows
    # или
    source venv/bin/activate      # Mac/Linux
3. Установи зависимости:
    ```bash
    pip install -r requirements.txt
## Использование
Выполняйте команды по порядку:
### Шаг 1: Парсинг оглавления
    ```bash
    python parse_toc.py

### Шаг 2: Парсинг задач
    ```bash
    python parse_tasks.py

### Шаг 3: Парсинг ответов
    ```bash
    python parse_answers.py

### Шаг 4: Объединение в итоговый файл
    ```bash
    python merge_all.py
## Результат
Итоговый файл `5_6_class_shevkin_merged.xlsx` будет содержать 3 листа:

|Лист|	Описание|
|----|----------|
tasks|	Задачи с ответами (объединённые)
author|	Информация об авторе
table_of_contents|	Оглавление с иерархией
## Формат итогового файла
Лист `tasks`
|Колонка|	Описание|
|----|----------|
id_tasks_book|	Номер задачи (например, 1.а), 6))
task|	Условие задачи
answer|	Ответ (из раздела "Ответы и советы")
classes|	Класс (всегда 5;6)
paragraph_id|	ID параграфа из оглавления
topic_id|	ID темы (всегда 1)
level|	Уровень сложности (всегда 1)

Лист `table_of_contents`
|Колонка|	Описание|
|----|----------|
id|	Уникальный идентификатор раздела
name|	Название раздела/подраздела
parent|	ID родительского раздела (0 для корневых)

Лист `author`
|Колонка|	Описание|
|----|----------|
name|	Полное название книги
author|	Автор
description|	Аннотация

## Описание скриптов
`parse_toc.py`

Парсит оглавление из документа Word и создаёт Excel-файл с иерархической структурой.

Вход: `input_data/tekstovye_zadachi_po_matematike.docx`

Выход: `output/результат_tekstovye_zadachi_po_matematike.xlsx`

`parse_tasks.py`

Парсит все задачи, распознавая номера, подпункты и условия. Автоматически определяет принадлежность задачи к параграфу на основе оглавления.

Вход: `input_data/tekstovye_zadachi_po_matematike.docx`

Выход: `output/задачи_tekstovye_zadachi_po_matematike.xlsx`

`parse_answers.py`

Парсит раздел "Ответы и советы", извлекая ответы для каждой задачи.

Вход: `input_data/tekstovye_zadachi_po_matematike.docx`

Выход: `output/ответы_tekstovye_zadachi_po_matematike.xlsx`

`merge_all.py`

Объединяет три промежуточных файла в один итоговый Excel с тремя листами.

Вход:
- output/задачи_*.xlsx
- output/ответы_*.xlsx
- output/результат_*.xlsx

Выход: `output/5_6_class_shevkin_merged.xlsx`

## Особенности
- Автоматическая привязка к параграфам: парсер считывает оглавление и автоматически определяет, к какому разделу относится каждая задача.
- Корректная обработка подпунктов: поддерживаются форматы 1.а), 1.1), 2) и т.д.
- Объединение задач с ответами: ответы автоматически сопоставляются с задачами по номеру.
- Устойчивость к ошибкам: парсер обрабатывает переносы строк, многострочные условия и специальные случаи.

## Возможные проблемы и решения
Закройте Excel-файл, который пытается перезаписать парсер.

Автор
[Максим Числавский](https://github.com/MaksimChislavskiy)