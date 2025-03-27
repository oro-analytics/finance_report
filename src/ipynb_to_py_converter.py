import nbformat
from nbconvert import PythonExporter
import os

# Путь к файлу
notebook_filename = "profit_center_analysis.ipynb"
# Другие настройки см поиском по слову "Настройки"

# Куда будет записан
python_filename = os.path.splitext(notebook_filename)[0] + ".py"

# Проверяем, существует ли файл
if os.path.exists(notebook_filename):
    # Открываем и загружаем ноутбук
    with open(notebook_filename, "r", encoding="utf-8") as nb_file:
        notebook_content = nbformat.read(nb_file, as_version=4)

    # Преобразуем в .py с помощью nbconvert
    python_exporter = PythonExporter()
    # Настройки
    python_exporter.exclude_input_prompt = True    # Убираем строки # In[номер строки]
    python_exporter.exclude_markdown = True        # Убираем ячейки с Markdown
    python_exporter.exclude_raw = True             # Убираем raw ячейки

    python_script, _ = python_exporter.from_notebook_node(notebook_content)

    # Сохраняем как .py файл
    with open(python_filename, "w", encoding="utf-8") as py_file:
        py_file.write(python_script)

    print(f"Файл '{notebook_filename}' успешно преобразован в '{python_filename}'")
else:
    print(f"Файл '{notebook_filename}' не найден в текущей папке.")