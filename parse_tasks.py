import os
import re
import openpyxl
from docx import Document

INPUT_FOLDER = 'input_data'
OUTPUT_FOLDER = 'output'
TARGET_FILE = 'tekstovye_zadachi_po_matematike.docx'

# Задача с буквой: "1. а)", "3. 1)", "207*. 2)"
TASK_WITH_LETTER_RE = re.compile(r'^(\d+\*?)\.\s*\S?\s*([а-я]\)|\d\))\s*')

# Продолжение задачи: "б)", "2)" (без номера задачи)
TASK_CONTINUE_RE = re.compile(r'^\s*\S?\s*([а-я]\)|\d\))\s*')

# Задача без буквы: "6. Первая бригада..."
TASK_NO_LETTER_RE = re.compile(r'^(\d+\*?)\.\s+\S?\s*([А-Я].+)')

# Раздел: 1-3 слова, без номера страницы
SECTION_IN_TASKS_RE = re.compile(
    r'^(\d+)\.\s+([А-Я][а-я]+(?:\s+[а-я]+){0,2})\s*$'
)

# Подраздел: цифра.цифра, любой текст
SUBSECTION_RE = re.compile(
    r'^(\d+\.\d+)\.?\s+(.+?)\s*$'
)

# Варианты через точку с запятой: "а) 1 рубля; б) 1 метра"
INLINE_VARIANTS_RE = re.compile(r';\s*[а-я\d]\)')

# Два и более подпункта в одной строке: "1) ... 2) ..."
MULTI_VARIANTS_RE = re.compile(r'[а-я\d]\).*[а-я\d]\)')

if TARGET_FILE:
    doc_path = os.path.join(INPUT_FOLDER, TARGET_FILE)
    if not os.path.exists(doc_path):
        raise FileNotFoundError(f'Файл {doc_path} не найден')
else:
    raise ValueError('Укажите TARGET_FILE')

os.makedirs(OUTPUT_FOLDER, exist_ok=True)

doc = Document(doc_path)
all_lines = []
for paragraph in doc.paragraphs:
    line = paragraph.text.strip()
    if line:
        all_lines.append(line)

tasks = []
current_task = None
current_task_num = None
mode = 'toc'

for line in all_lines:
    # === Смена режимов ===
    if line == 'Оглавление' and mode != 'answers':
        mode = 'toc'
        continue

    if line.startswith('Ответы и советы'):
        mode = 'answers'
        continue

    if line == 'Оглавление' and mode == 'answers':
        break

    # === Парсинг содержания (пропускаем) ===
    if mode == 'toc':
        if SECTION_IN_TASKS_RE.match(line) or SUBSECTION_RE.match(line):
            continue
        else:
            mode = 'tasks'

    # === Парсинг задач ===
    if mode == 'tasks':
        # Пропускаем заголовки разделов и подразделов
        if SECTION_IN_TASKS_RE.match(line):
            continue
        if SUBSECTION_RE.match(line):
            continue

        full_match = TASK_WITH_LETTER_RE.match(line)
        continue_match = TASK_CONTINUE_RE.match(line)
        no_letter_match = TASK_NO_LETTER_RE.match(line)

        if full_match:
            text_after = line[full_match.end():].strip()

            # Если в строке варианты — это продолжение
            if INLINE_VARIANTS_RE.search(text_after) or MULTI_VARIANTS_RE.search(text_after):
                if current_task:
                    current_task['task'] = current_task['task'] + " " + line
                continue

            # Сохраняем предыдущую задачу, если она не заканчивается на ":"
            if current_task and not current_task['task'].rstrip().endswith(':'):
                current_task['task'] = current_task['task'].strip()
                tasks.append(current_task)
            elif current_task:
                # Заканчивается на ":" — дополняем, а не сохраняем
                current_task['task'] = current_task['task'] + " " + line
                # Обновляем номер задачи, если изменился
                current_task_num = full_match.group(1)
                current_task['id_tasks_book'] = f"{current_task_num}{full_match.group(2)}"
                continue

            current_task_num = full_match.group(1)
            sub_letter = full_match.group(2)
            current_task = {
                'id_tasks_book': f"{current_task_num}{sub_letter}",
                'task': text_after,
                'classes': '5,6',
                'level': '1',
            }

        elif continue_match and current_task_num:
            text_after = line[continue_match.end():].strip()

            # Если в строке варианты — продолжение
            if INLINE_VARIANTS_RE.search(text_after) or MULTI_VARIANTS_RE.search(text_after):
                if current_task:
                    current_task['task'] = current_task['task'] + " " + line
                continue

            # Сохраняем предыдущую задачу, если она не заканчивается на ":"
            if current_task and not current_task['task'].rstrip().endswith(':'):
                current_task['task'] = current_task['task'].strip()
                tasks.append(current_task)
            elif current_task:
                current_task['task'] = current_task['task'] + " " + line
                continue

            sub_letter = continue_match.group(1)
            current_task = {
                'id_tasks_book': f"{current_task_num}{sub_letter}",
                'task': text_after,
                'classes': '5,6',
                'level': '1',
            }

        elif no_letter_match:
            # Сохраняем предыдущую задачу, если она не заканчивается на ":"
            if current_task and not current_task['task'].rstrip().endswith(':'):
                current_task['task'] = current_task['task'].strip()
                tasks.append(current_task)
            elif current_task:
                current_task['task'] = current_task['task'] + " " + line
                continue

            current_task_num = no_letter_match.group(1)
            text = no_letter_match.group(2).strip()
            current_task = {
                'id_tasks_book': current_task_num,
                'task': text,
                'classes': '5,6',
                'level': '1',
            }

        else:
            # Продолжение текста задачи
            if current_task:
                current_task['task'] = current_task['task'] + " " + line

# Сохраняем последнюю задачу
if current_task:
    current_task['task'] = current_task['task'].strip()
    tasks.append(current_task)

# === Excel ===
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "tasks"
ws.append(["id_tasks_book", "task", "answer", "classes", "topic_id", "level", "paragraph_id"])

for task in tasks:
    ws.append([
        task.get('id_tasks_book', ''),
        task.get('task', ''),
        task.get('answer', ''),
        task.get('classes', '5,6'),
        task.get('topic_id', ''),
        task.get('level', '1'),
        task.get('paragraph_id', ''),
    ])

output_name = f"задачи_{TARGET_FILE.replace('.docx', '.xlsx')}"
output_path = os.path.join(OUTPUT_FOLDER, output_name)

try:
    wb.save(output_path)
    print(f"Готово! {len(tasks)} задач сохранено в {output_path}")
except PermissionError:
    print(f"Ошибка: закройте файл {output_path} перед повторным запуском")