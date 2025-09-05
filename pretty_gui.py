import os
import re
import sys
import shutil
import threading
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import pythoncom  # Инициализация COM
from datetime import datetime
from docxtpl import DocxTemplate
from docx2pdf import convert
from openpyxl import load_workbook
from PyPDF2 import PdfReader, PdfWriter
import fitz
import time
from tqdm import tqdm
from PIL import Image


def is_filled(value):
    """
    Проверяет, что значение не является пустым.
    Возвращает False для None, пустых строк и строк из одних пробелов.
    """
    if value is None:
        return False
    if isinstance(value, str) and value.strip() == "":
        return False
    return True


def image_to_pdf(image_path, pdf_path, a4_size=(595, 842)):
    """
    Конвертирует изображение в PDF с правильным масштабированием под A4
    """
    try:
        print(f"🖼️ Конвертация изображения в PDF: {os.path.basename(image_path)}")

        # Открываем изображение
        img = Image.open(image_path)

        # Конвертируем в RGB если необходимо (для PNG с прозрачностью)
        if img.mode in ('RGBA', 'LA', 'P'):
            # Создаем белый фон
            background = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
            img = background
        elif img.mode != 'RGB':
            img = img.convert('RGB')

        img_width, img_height = img.size

        # Вычисляем масштаб для вписывания в A4 с сохранением пропорций
        scale_x = a4_size[0] / img_width
        scale_y = a4_size[1] / img_height
        scale = min(scale_x, scale_y)

        # Новые размеры с учетом масштаба
        new_width = int(img_width * scale)
        new_height = int(img_height * scale)

        # Создаем PDF документ
        pdf_doc = fitz.open()
        page = pdf_doc.new_page(width=a4_size[0], height=a4_size[1])

        # Центрируем изображение на странице
        x_offset = (a4_size[0] - new_width) / 2
        y_offset = (a4_size[1] - new_height) / 2

        # Создаем прямоугольник для размещения изображения
        rect = fitz.Rect(x_offset, y_offset, x_offset + new_width, y_offset + new_height)

        # Вставляем изображение
        page.insert_image(rect, filename=image_path)

        # Сохраняем PDF
        pdf_doc.save(pdf_path)
        pdf_doc.close()

        print(f"   ✅ Изображение конвертировано в PDF")

    except Exception as e:
        print(f"   ❌ Ошибка конвертации изображения {image_path}: {str(e)}")
        raise


class ToolTip:
    """Класс для создания всплывающих подсказок"""

    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.on_enter)
        self.widget.bind("<Leave>", self.on_leave)
        self.tooltip = None

    def on_enter(self, event=None):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 20
        y += self.widget.winfo_rooty() + 20

        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")

        label = tk.Label(self.tooltip, text=self.text,
                         background="lightyellow",
                         relief="solid", borderwidth=1,
                         font=("Arial", "8", "normal"))
        label.pack()

    def on_leave(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None


def choose_files_and_folders(parent, callback):
    """
    Улучшенное окно для выбора входных данных с исправленной валидацией
    """
    root = tk.Toplevel(parent)
    root.title("🔧 Настройка параметров обработки документов")
    root.geometry("1000x900")
    root.resizable(True, True)

    # Настройка стилей
    style = ttk.Style()
    style.theme_use('clam')

    # Переменные для хранения путей и флагов
    passports_folder = tk.StringVar()
    lab_folder = tk.StringVar()
    executive_folder = tk.StringVar()
    output_folder = tk.StringVar()
    excel_file = tk.StringVar()
    word_template = tk.StringVar()
    double_sided_print = tk.BooleanVar(value=True)
    black_and_white = tk.BooleanVar()

    # Переменные для валидации
    validation_vars = {
        'passports': tk.BooleanVar(),
        'lab': tk.BooleanVar(),
        'executive': tk.BooleanVar(),
        'output': tk.BooleanVar(),
        'excel': tk.BooleanVar(),
        'word': tk.BooleanVar()
    }

    # Главный заголовок
    header_frame = ttk.Frame(root)
    header_frame.pack(fill=tk.X, padx=20, pady=(20, 10))

    title_label = ttk.Label(header_frame, text="📄 Система автоматической обработки документов",
                            font=("Arial", 16, "bold"))
    title_label.pack()

    subtitle_label = ttk.Label(header_frame, text="Выберите папки и файлы для обработки",
                               font=("Arial", 10))
    subtitle_label.pack()

    # Создаем Notebook для вкладок
    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

    # Вкладка 1: Основные настройки
    main_frame = ttk.Frame(notebook)
    notebook.add(main_frame, text="📁 Основные настройки")

    def on_drop(event, var, validation_var=None):
        file_path = event.data.strip('{}')
        if is_filled(file_path) and os.path.exists(file_path):
            var.set(file_path)
            if validation_var:
                validation_var.set(True)
                update_submit_button()

    def validate_path(var, validation_var, is_file=False):
        """Проверяет существование пути и обновляет индикатор с улучшенной валидацией"""
        path = var.get()

        # Проверяем, что путь не пустой и не состоит из пробелов
        if not is_filled(path):
            validation_var.set(False)
            update_submit_button()
            return

        # Проверяем существование пути
        if os.path.exists(path):
            if is_file and os.path.isfile(path):
                validation_var.set(True)
            elif not is_file and os.path.isdir(path):
                validation_var.set(True)
            else:
                validation_var.set(False)
        else:
            validation_var.set(False)

        update_submit_button()

    def create_path_section(parent, title, description, variable, select_type,
                            validation_var, filetypes=None, tooltip_text=""):
        """Создает секцию для выбора пути с улучшенным дизайном"""

        # Основная рамка секции
        section_frame = ttk.LabelFrame(parent, text=title, padding=(10, 5))
        section_frame.pack(fill=tk.X, padx=10, pady=8)

        # Описание
        if description:
            desc_label = ttk.Label(section_frame, text=description,
                                   font=("Arial", 9), foreground="gray")
            desc_label.pack(anchor="w")

        # Рамка для поля ввода и кнопок
        input_frame = ttk.Frame(section_frame)
        input_frame.pack(fill=tk.X, pady=(5, 0))

        # Поле ввода
        entry = ttk.Entry(input_frame, textvariable=variable, font=("Arial", 9))
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Индикатор валидации
        status_label = ttk.Label(input_frame, text="❌", font=("Arial", 12))
        status_label.pack(side=tk.LEFT, padx=(5, 0))

        # Кнопка обзора
        def browse():
            if select_type == 'folder':
                path = filedialog.askdirectory(parent=parent, title=title)
            else:
                path = filedialog.askopenfilename(
                    parent=parent, title=title,
                    filetypes=filetypes or [('Все файлы', '*.*')]
                )
            if is_filled(path):  # Проверяем на заполненность
                variable.set(path)
                validate_path(variable, validation_var, select_type == 'file')

        browse_btn = ttk.Button(input_frame, text="📁 Обзор", command=browse)
        browse_btn.pack(side=tk.LEFT, padx=(5, 0))

        # Поддержка Drag & Drop
        entry.drop_target_register(DND_FILES)
        entry.dnd_bind('<<Drop>>', lambda e: on_drop(e, variable, validation_var))

        # Валидация при изменении текста
        def on_change(*args):
            validate_path(variable, validation_var, select_type == 'file')

        variable.trace('w', on_change)

        # Обновление индикатора
        def update_status():
            if validation_var.get():
                status_label.config(text="✅", foreground="green")
            else:
                status_label.config(text="❌", foreground="red")

        validation_var.trace('w', lambda *args: update_status())

        # Подсказка
        if tooltip_text:
            ToolTip(entry, tooltip_text)

        return section_frame

    # Создаем секции
    create_path_section(main_frame, "📋 Папка с паспортами",
                        "Выберите папку, содержащую файлы паспортов материалов (PDF, Word, изображения)",
                        passports_folder, 'folder', validation_vars['passports'],
                        tooltip_text="Поддерживаются: PDF, DOCX, DOC, JPG, PNG, BMP, TIFF")

    create_path_section(main_frame, "🔬 Папка с лабораторными заключениями",
                        "Выберите папку с лабораторными заключениями и сертификатами (PDF, Word, изображения)",
                        lab_folder, 'folder', validation_vars['lab'],
                        tooltip_text="Поддерживаются: PDF, DOCX, DOC, JPG, PNG, BMP, TIFF")

    create_path_section(main_frame, "📐 Папка с исполнительными схемами",
                        "Выберите папку с исполнительными схемами и чертежами (PDF, Word, изображения)",
                        executive_folder, 'folder', validation_vars['executive'],
                        tooltip_text="Поддерживаются: PDF, DOCX, DOC, JPG, PNG, BMP, TIFF")

    create_path_section(main_frame, "💾 Папка для результата",
                        "Выберите папку, куда сохранить обработанные документы",
                        output_folder, 'folder', validation_vars['output'],
                        tooltip_text="В эту папку будут сохранены все обработанные файлы")

    create_path_section(main_frame, "📊 Excel-файл с данными",
                        "Файл Excel с таблицами 'Реквизиты' и 'АСР ТАБЛ'",
                        excel_file, 'file', validation_vars['excel'],
                        filetypes=[('Excel файлы', '*.xlsx *.xlsm'), ('Все файлы', '*.*')],
                        tooltip_text="Файл должен содержать листы 'Реквизиты' и 'АСР ТАБЛ'")

    create_path_section(main_frame, "📝 Шаблон Word документа",
                        "Шаблон Word для создания АОСР документов",
                        word_template, 'file', validation_vars['word'],
                        filetypes=[('Документы Word', '*.docx'), ('Все файлы', '*.*')],
                        tooltip_text="Шаблон должен содержать переменные для замещения")

    # Вкладка 2: Дополнительные настройки
    options_frame = ttk.Frame(notebook)
    notebook.add(options_frame, text="⚙️ Дополнительные настройки")

    # Настройки PDF
    pdf_frame = ttk.LabelFrame(options_frame, text="🖨️ Настройки PDF", padding=(15, 10))
    pdf_frame.pack(fill=tk.X, padx=20, pady=20)

    double_sided_cb = ttk.Checkbutton(pdf_frame, text="Двусторонняя печать PDF",
                                      variable=double_sided_print)
    double_sided_cb.pack(anchor="w", pady=5)
    ToolTip(double_sided_cb, "Добавляет пустые страницы для правильной двусторонней печати")

    black_white_cb = ttk.Checkbutton(pdf_frame, text="Конвертировать PDF в черно-белый",
                                     variable=black_and_white)
    black_white_cb.pack(anchor="w", pady=5)
    ToolTip(black_white_cb, "Преобразует цветные PDF в черно-белые для экономии чернил")

    # Информационная панель
    info_frame = ttk.LabelFrame(options_frame, text="ℹ️ Информация", padding=(15, 10))
    info_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

    info_text = tk.Text(info_frame, height=10, wrap=tk.WORD, font=("Arial", 9))
    info_text.pack(fill=tk.BOTH, expand=True)

    info_content = """Программа выполняет следующие действия:

1. 📄 Загружает данные из Excel файла и Word шаблона
2. 🔄 Для каждой строки в таблице создает заполненный АОСР документ
3. 📁 Копирует соответствующие паспорта, лабораторные заключения и схемы
4. 🖼️ Изображения автоматически конвертируются в PDF с правильным масштабированием
5. 🖨️ Применяет настройки печати (двусторонняя печать, черно-белый режим)
6. 📋 Объединяет все файлы в один итоговый PDF документ

📝 Поддерживаемые форматы:
• Документы: PDF, DOCX, DOC  
• Изображения: JPG, JPEG, PNG, BMP, TIFF

💡 Советы:
• Убедитесь, что Excel файл содержит листы 'Реквизиты' и 'АСР ТАБЛ'
• Имена файлов в Excel должны соответствовать файлам в папках (с расширением или без)
• Используйте точку с запятой (;) для разделения нескольких файлов
• Изображения будут масштабированы под формат A4 с сохранением пропорций"""

    info_text.insert(tk.END, info_content)
    info_text.config(state=tk.DISABLED)

    # Нижняя панель с кнопками
    bottom_frame = ttk.Frame(root)
    bottom_frame.pack(fill=tk.X, padx=20, pady=(0, 20))

    # Индикатор готовности
    status_frame = ttk.Frame(bottom_frame)
    status_frame.pack(side=tk.LEFT)

    status_label = ttk.Label(status_frame, text="📋 Статус: Заполните все поля",
                             font=("Arial", 10))
    status_label.pack()

    progress_label = ttk.Label(status_frame, text="", font=("Arial", 9), foreground="gray")
    progress_label.pack()

    # Кнопки
    button_frame = ttk.Frame(bottom_frame)
    button_frame.pack(side=tk.RIGHT)

    def update_submit_button():
        """Обновляет состояние кнопки подтверждения и статус с улучшенной проверкой"""
        all_valid = all(var.get() for var in validation_vars.values())

        if all_valid:
            submit_btn.config(state='normal')
            status_label.config(text="✅ Готов к обработке", foreground="green")
            progress_label.config(text="Все поля заполнены корректно")
        else:
            submit_btn.config(state='disabled')
            missing = [name for name, var in validation_vars.items() if not var.get()]
            status_label.config(text="⚠️ Заполните все поля", foreground="orange")
            progress_label.config(text=f"Не заполнено: {', '.join(missing)}")

    result = []

    def on_submit():
        # Дополнительная проверка с использованием is_filled
        paths_to_check = [
            (passports_folder.get(), "папка с паспортами"),
            (lab_folder.get(), "папка с лабораторными заключениями"),
            (executive_folder.get(), "папка с исполнительными схемами"),
            (output_folder.get(), "папка для результата"),
            (excel_file.get(), "Excel-файл"),
            (word_template.get(), "шаблон Word")
        ]

        empty_fields = []
        for path, description in paths_to_check:
            if not is_filled(path):
                empty_fields.append(description)

        if empty_fields:
            messagebox.showwarning("Предупреждение",
                                   f"Следующие поля не заполнены или содержат только пробелы:\n\n" +
                                   "\n".join(f"• {field}" for field in empty_fields),
                                   parent=root)
            return

        # Проверяем существование путей
        non_existing = []
        for path, description in paths_to_check:
            if not os.path.exists(path.strip()):
                non_existing.append(f"{description}: {path}")

        if non_existing:
            messagebox.showerror("Ошибка",
                                 f"Следующие пути не существуют:\n\n" +
                                 "\n".join(f"• {item}" for item in non_existing),
                                 parent=root)
            return

        result.extend([
            passports_folder.get().strip(),
            lab_folder.get().strip(),
            executive_folder.get().strip(),
            output_folder.get().strip(),
            excel_file.get().strip(),
            word_template.get().strip(),
            double_sided_print.get(),
            black_and_white.get()
        ])
        root.destroy()

    def on_close():
        callback()
        root.destroy()

    submit_btn = ttk.Button(button_frame, text="🚀 Начать обработку",
                            command=on_submit, state='disabled')
    submit_btn.pack(side=tk.LEFT, padx=(10, 5))

    cancel_btn = ttk.Button(button_frame, text="❌ Отмена", command=on_close)
    cancel_btn.pack(side=tk.LEFT, padx=5)

    # Привязка горячих клавиш
    root.bind('<Return>', lambda e: on_submit() if submit_btn['state'] == 'normal' else None)
    root.bind('<Escape>', lambda e: on_close())

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.grab_set()

    # Центрирование окна
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")

    parent.wait_window(root)

    if not result:
        return None
    return tuple(result)


def clear_output_folder(output_folder):
    """Очищает папку вывода, перемещая старые файлы в архив"""
    archive_folder = os.path.join(output_folder, 'архив')
    if not os.path.exists(archive_folder):
        os.makedirs(archive_folder)

    moved_files = 0
    for file_name in os.listdir(output_folder):
        file_path = os.path.join(output_folder, file_name)
        if os.path.isfile(file_path):
            shutil.move(file_path, os.path.join(archive_folder, file_name))
            moved_files += 1

    if moved_files > 0:
        print(f'📦 Перемещено {moved_files} старых файлов в архивную папку.')


def add_blank_pages(file_path):
    """Добавляет пустую страницу после каждой страницы PDF для двусторонней печати"""
    print(f"🖨️ Подготовка к двусторонней печати: {os.path.basename(file_path)}")
    reader = PdfReader(file_path)
    writer = PdfWriter()

    for i, page in enumerate(reader.pages):
        writer.add_page(page)
        writer.add_blank_page()
        if (i + 1) % 10 == 0:  # Прогресс каждые 10 страниц
            print(f"   Обработано страниц: {i + 1}/{len(reader.pages)}")

    with open(file_path, 'wb') as output_file:
        writer.write(output_file)
    print(f"✅ Добавлено {len(reader.pages)} пустых страниц")


def convert_to_black_and_white(file_path):
    """Конвертирует PDF в черно-белый и масштабирует под A4"""
    print(f"🎨 Конвертация в черно-белый: {os.path.basename(file_path)}")
    doc = fitz.open(file_path)
    a4_width = 595
    a4_height = 842

    temp_file = file_path + ".temp"
    temp_doc = fitz.open()

    for page_number, page in enumerate(doc, start=1):
        if page_number % 5 == 0:  # Прогресс каждые 5 страниц
            print(f"   Обработка страницы {page_number} из {len(doc)}...")

        pix = page.get_pixmap(dpi=150)
        pix = fitz.Pixmap(fitz.csGRAY, pix)

        img_width, img_height = pix.width, pix.height
        scale_x = a4_width / img_width
        scale_y = a4_height / img_height
        scale = min(scale_x, scale_y)

        scaled_width = img_width * scale
        scaled_height = img_height * scale
        x_offset = (a4_width - scaled_width) / 2
        y_offset = (a4_height - scaled_height) / 2

        new_page = temp_doc.new_page(width=a4_width, height=a4_height)
        new_rect = fitz.Rect(x_offset, y_offset, x_offset + scaled_width, y_offset + scaled_height)
        new_page.insert_image(new_rect, stream=pix.tobytes("png"))

    temp_doc.save(temp_file, garbage=3, deflate=True)
    temp_doc.close()
    doc.close()
    os.replace(temp_file, file_path)
    print("✅ Конвертация завершена")


def copy_and_rename_files(source_folder, file_names_str, start_index, prefix,
                          double_sided_print, black_and_white, output_folder):
    """Копирует и переименовывает файлы с заданным префиксом, поддерживая документы и изображения"""
    new_files = []

    # Поддерживаемые расширения
    document_extensions = ['.docx', '.doc', '.pdf']
    image_extensions = ['.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif']
    all_extensions = document_extensions + image_extensions

    # Исправленная проверка заполненности строки с именами файлов
    if not is_filled(file_names_str):
        return new_files, start_index

    # Фильтруем пустые имена файлов и имена из одних пробелов
    file_names = [name.strip() for name in file_names_str.split(';') if is_filled(name)]

    if not file_names:  # Если после фильтрации не осталось файлов
        return new_files, start_index

    max_retries = 3
    retry_delay = 2

    print(f"📁 Копирование файлов категории '{prefix}' ({len(file_names)} файлов)")

    for file_name in file_names:
        if is_filled(file_name):
            found = False

            # Проверяем, есть ли уже расширение в имени файла
            file_name_lower = file_name.lower()
            has_extension = any(file_name_lower.endswith(ext) for ext in all_extensions)

            if has_extension:
                # Если расширение уже есть, ищем файл как есть
                source_path = os.path.join(source_folder, file_name)
                if os.path.exists(source_path):
                    # Извлекаем расширение для нового имени
                    name_without_ext = os.path.splitext(file_name)[0]
                    extension = os.path.splitext(file_name)[1]

                    new_file_name = f"{start_index:03d}_{prefix}_{name_without_ext}{extension}"
                    dest_path = os.path.join(output_folder, new_file_name)

                    for attempt in range(max_retries):
                        try:
                            shutil.copy2(source_path, dest_path)
                            print(f"   ✅ {new_file_name}")

                            # Обработка изображений - конвертируем в PDF
                            if extension.lower() in image_extensions:
                                pdf_dest_path = dest_path.replace(extension, '.pdf')
                                image_to_pdf(dest_path, pdf_dest_path)
                                os.remove(dest_path)  # Удаляем оригинальное изображение
                                dest_path = pdf_dest_path
                                new_file_name = new_file_name.replace(extension, '.pdf')

                            # Обработка PDF
                            if dest_path.lower().endswith('.pdf'):
                                if black_and_white:
                                    convert_to_black_and_white(dest_path)
                                if double_sided_print:
                                    add_blank_pages(dest_path)

                            new_files.append(os.path.basename(dest_path))
                            start_index += 1
                            found = True
                            break

                        except (PermissionError, OSError) as e:
                            if attempt < max_retries - 1:
                                print(f"   ⚠️ Ошибка при копировании {file_name}. Повтор через {retry_delay} сек...")
                                time.sleep(retry_delay)
                            else:
                                print(f"   ❌ Не удалось скопировать {file_name}: {str(e)}")
                else:
                    print(f"   ❌ Файл не найден: {file_name}")
            else:
                # Если расширения нет, пробуем добавить все поддерживаемые расширения
                for ext in all_extensions:
                    source_path = os.path.join(source_folder, file_name + ext)
                    if os.path.exists(source_path):
                        new_file_name = f"{start_index:03d}_{prefix}_{file_name}{ext}"
                        dest_path = os.path.join(output_folder, new_file_name)

                        for attempt in range(max_retries):
                            try:
                                shutil.copy2(source_path, dest_path)
                                print(f"   ✅ {new_file_name}")

                                # Обработка изображений - конвертируем в PDF
                                if ext.lower() in image_extensions:
                                    pdf_dest_path = dest_path.replace(ext, '.pdf')
                                    image_to_pdf(dest_path, pdf_dest_path)
                                    os.remove(dest_path)  # Удаляем оригинальное изображение
                                    dest_path = pdf_dest_path
                                    new_file_name = new_file_name.replace(ext, '.pdf')

                                # Обработка PDF
                                if dest_path.lower().endswith('.pdf'):
                                    if black_and_white:
                                        convert_to_black_and_white(dest_path)
                                    if double_sided_print:
                                        add_blank_pages(dest_path)

                                new_files.append(os.path.basename(dest_path))
                                start_index += 1
                                found = True
                                break

                            except (PermissionError, OSError) as e:
                                if attempt < max_retries - 1:
                                    print(
                                        f"   ⚠️ Ошибка при копировании {file_name}{ext}. Повтор через {retry_delay} сек...")
                                    time.sleep(retry_delay)
                                else:
                                    print(f"   ❌ Не удалось скопировать {file_name}{ext}: {str(e)}")
                        break  # Выходим из цикла расширений, если файл найден

                if not found:
                    extensions_list = ", ".join(all_extensions)
                    print(f"   ❌ Файл не найден: {file_name} (проверены расширения: {extensions_list})")

    return new_files, start_index


def merge_output_files(output_folder):
    """Объединяет все файлы из папки вывода в один PDF (включая изображения)"""
    print("📑 Начало объединения файлов (PDF, DOCX и изображения)...")

    # Поддерживаемые форматы для объединения
    supported_extensions = ('.pdf', '.docx')
    image_extensions = ('.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif')

    # Получаем все файлы для объединения
    all_files = sorted([f for f in os.listdir(output_folder)
                        if (f.lower().endswith(supported_extensions + image_extensions))
                        and not f.startswith('объединенный')])

    if not all_files:
        print("⚠️ Нет файлов для объединения")
        return

    merged_pdf = fitz.open()

    with tqdm(total=len(all_files), desc="📄 Объединение", unit="файл",
              bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}]") as pbar:

        for file in all_files:
            file_path = os.path.join(output_folder, file)

            try:
                if file.lower().endswith('.pdf'):
                    # Обычный PDF файл
                    pdf_document = fitz.open(file_path)
                    merged_pdf.insert_pdf(pdf_document)
                    pdf_document.close()

                elif file.lower().endswith('.docx'):
                    # Word документ - конвертируем в PDF
                    print(f"   🔄 Конвертация Word: {file}")
                    temp_pdf = file_path.replace('.docx', '_temp.pdf')
                    convert(file_path, temp_pdf)
                    pdf_document = fitz.open(temp_pdf)
                    merged_pdf.insert_pdf(pdf_document)
                    pdf_document.close()
                    os.remove(temp_pdf)

                elif any(file.lower().endswith(ext) for ext in image_extensions):
                    # Изображение - конвертируем в PDF
                    print(f"   🖼️ Конвертация изображения: {file}")
                    temp_pdf = file_path + '_temp.pdf'
                    image_to_pdf(file_path, temp_pdf)
                    pdf_document = fitz.open(temp_pdf)
                    merged_pdf.insert_pdf(pdf_document)
                    pdf_document.close()
                    os.remove(temp_pdf)

                pbar.set_postfix_str(f"Обработан: {file[:30]}...")
                pbar.update(1)

            except Exception as e:
                print(f"   ❌ Ошибка обработки {file}: {str(e)}")
                pbar.update(1)

    merged_file_path = os.path.join(output_folder, "объединенный_документ.pdf")
    merged_pdf.save(merged_file_path)
    merged_pdf.close()
    print(f"🎉 Объединенный документ сохранен: {os.path.basename(merged_file_path)}")


def run_processing(passports_folder, lab_folder, executive_folder,
                   output_folder, excel_file, word_template,
                   double_sided_print, black_and_white):
    """Основная функция обработки файлов с поддержкой изображений"""
    print("🚀 === НАЧАЛО ОБРАБОТКИ ДОКУМЕНТОВ ===")
    print(f"📅 Время запуска: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}")
    print("=" * 50)

    max_retries = 3
    retry_delay = 2

    try:
        # Подготовка папки вывода
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            print(f"📁 Создана папка вывода: {output_folder}")

        clear_output_folder(output_folder)

        # Загрузка шаблона Word
        print("\n📝 ЗАГРУЗКА ШАБЛОНА WORD")
        print("-" * 30)
        for attempt in range(max_retries):
            try:
                print(f"   Загрузка: {os.path.basename(word_template)}")
                doc_template = DocxTemplate(word_template)
                print("   ✅ Шаблон Word успешно загружен")
                break
            except Exception as e:
                if attempt < max_retries - 1:
                    print(f"   ⚠️ Ошибка загрузки. Повтор через {retry_delay} сек...")
                    time.sleep(retry_delay)
                else:
                    raise Exception(f"Не удалось загрузить шаблон Word: {str(e)}")

        # Загрузка Excel
        print("\n📊 ЗАГРУЗКА EXCEL ФАЙЛА")
        print("-" * 30)
        for attempt in range(max_retries):
            try:
                print(f"   Загрузка: {os.path.basename(excel_file)}")
                wb = load_workbook(excel_file, data_only=True)
                print("   ✅ Excel файл успешно загружен")
                break
            except Exception as e:
                if attempt < max_retries - 1:
                    print(f"   ⚠️ Ошибка загрузки. Повтор через {retry_delay} сек...")
                    time.sleep(retry_delay)
                else:
                    raise Exception(f"Не удалось загрузить Excel файл: {str(e)}")

        # Загрузка листов
        try:
            sheet_requisites = wb["Реквизиты"]
            sheet_asr = wb['АСР ТАБЛ']
            print("   ✅ Листы 'Реквизиты' и 'АСР ТАБЛ' найдены")
        except KeyError as e:
            raise Exception(f"Отсутствует лист {e} в Excel файле")

        # Настройка столбцов
        EXCEL_COLUMNS = {
            'Номер_акта': 'Номер акта',
            'Имя_работы': 'Наименование работы',
            'Объем': 'Объем',
            'Ед_изм': 'Ед.изм.',
            'Начало_работ': 'Дата начала',
            'Конец_работ': 'Дата конца',
            'Дата_составления_акта': 'Дата составления акта',
            'Материалы': 'Примененные материалы',
            'Последующие_работы': 'последующие работы',
            'Проект': 'Проект',
            'Лабы_по_материалам': 'Лабораторные заключениям',
            'Паспорта_по_материалам': 'Паспорта по материалам',
            'Схема': 'Исполнительные схемы',
            'Паспорта_файлы': 'Имя файлов паспортов',
            'Исполнительные_схемы': 'Имя файлов схем',
            'Лабораторные_файлы': 'Имя файлов лаб'
        }

        def find_column_indices(sheet, column_names):
            """Поиск индексов столбцов с подробным выводом"""
            column_indices = {}
            header_row = list(sheet.iter_rows(min_row=1, max_row=1, values_only=True))[0]

            print(f"\n🔍 АНАЛИЗ СТРУКТУРЫ EXCEL")
            print("-" * 30)
            print(f"   Найдено столбцов: {len([h for h in header_row if is_filled(h)])}")

            for var_name, excel_name in column_names.items():
                found_index = None
                for i, cell_value in enumerate(header_row):
                    if is_filled(cell_value) and str(cell_value).strip().lower() == excel_name.lower():
                        found_index = i
                        break

                if found_index is not None:
                    column_indices[var_name] = found_index
                    print(f"   ✅ '{excel_name}' -> колонка {found_index}")
                else:
                    column_indices[var_name] = None
                    print(f"   ❌ '{excel_name}' не найден")

            return column_indices

        column_indices = find_column_indices(sheet_asr, EXCEL_COLUMNS)

        # Проверка критических столбцов
        critical_columns = ['Номер_акта', 'Имя_работы']
        missing_critical = [col for col in critical_columns if column_indices.get(col) is None]
        if missing_critical:
            raise Exception(f"Отсутствуют критические столбцы: {missing_critical}")

        # Настройка форматирования дат
        MONTHS_RU = {
            1: 'января', 2: 'февраля', 3: 'марта', 4: 'апреля',
            5: 'мая', 6: 'июня', 7: 'июля', 8: 'августа',
            9: 'сентября', 10: 'октября', 11: 'ноября', 12: 'декабря'
        }

        def format_date(date_obj):
            if isinstance(date_obj, datetime):
                return f"{date_obj.day} {MONTHS_RU[date_obj.month]} {date_obj.year} г."
            return date_obj

        def process_value(value):
            if not is_filled(value):
                return ""
            if isinstance(value, str) and re.match(r'^[A-Z]+\d+$', value.strip()):
                try:
                    cell_value = sheet_requisites[value.strip()].value
                    if not is_filled(cell_value):
                        print(f"   ⚠️ Пустая ячейка: {value}")
                        return ""
                    return format_date(cell_value)
                except:
                    print(f"   ❌ Ошибка чтения ячейки: {value}")
                    return ""
            return value

        # Обработка реквизитов
        print(f"\n⚙️ ПОДГОТОВКА РЕКВИЗИТОВ")
        print("-" * 30)
        context_requisites = {}
        template_vars = doc_template.get_undeclared_template_variables()

        for key in template_vars:
            context_requisites[key] = process_value(key)

        print(f"   ✅ Обработано переменных: {len(context_requisites)}")

        # Подсчет строк для обработки с улучшенной проверкой
        total_rows = 0
        work_name_col_index = column_indices.get('Имя_работы', 1)

        for row in sheet_asr.iter_rows(min_row=2):
            if len(row) > work_name_col_index and is_filled(row[work_name_col_index].value):
                total_rows += 1
            else:
                break

        print(f"\n📋 ОБРАБОТКА СТРОК ДАННЫХ")
        print("-" * 30)
        print(f"   Всего строк к обработке: {total_rows}")

        # Основной цикл обработки
        file_index = 1
        processed_rows = 0

        for row_num, row in enumerate(sheet_asr.iter_rows(min_row=2, values_only=True), start=2):
            if not row or len(row) == 0:
                break

            # Улучшенная проверка наличия данных в строке
            work_name_col = column_indices.get('Имя_работы')
            if work_name_col is not None and len(row) > work_name_col:
                if not is_filled(row[work_name_col]):
                    break
            else:
                break

            processed_rows += 1
            work_name = str(row[work_name_col]).strip() if is_filled(row[work_name_col]) else "Без названия"
            print(f"\n   📄 Строка {processed_rows}/{total_rows}: {work_name}")

            # Извлечение данных строки с улучшенной обработкой
            row_data = {}
            for var_name, col_index in column_indices.items():
                try:
                    if col_index is not None and col_index < len(row):
                        val = row[col_index]
                    else:
                        val = None
                except IndexError:
                    val = None

                # Обработка значения с проверкой на заполненность
                if var_name in ['Начало_работ', 'Конец_работ', 'Дата_составления_акта']:
                    row_data[var_name] = format_date(val) if is_filled(val) else ""
                else:
                    row_data[var_name] = str(val).strip() if is_filled(val) else ""

            # Создание контекста и рендеринг
            context = {**context_requisites, **row_data}
            context = {k: (v if is_filled(v) else "") for k, v in context.items()}

            doc_template.render(context)
            aosr_filename = f"{file_index:03d}_АОСР_заполненный.docx"
            doc_template.save(os.path.join(output_folder, aosr_filename))
            print(f"      ✅ Создан документ: {aosr_filename}")
            file_index += 1

            # Копирование файлов с поддержкой изображений
            passport_files, file_index = copy_and_rename_files(
                passports_folder, row_data.get('Паспорта_файлы', ''),
                file_index, 'Паспорт', double_sided_print, black_and_white, output_folder
            )

            lab_files, file_index = copy_and_rename_files(
                lab_folder, row_data.get('Лабораторные_файлы', ''),
                file_index, 'Лаборатория', double_sided_print, black_and_white, output_folder
            )

            exec_files, file_index = copy_and_rename_files(
                executive_folder, row_data.get('Исполнительные_схемы', ''),
                file_index, 'Исполнительная_схема', double_sided_print, False, output_folder
            )

            total_copied = len(passport_files) + len(lab_files) + len(exec_files)
            print(f"      📎 Скопировано файлов: {total_copied}")

        # Финальное объединение
        print(f"\n🔗 ОБЪЕДИНЕНИЕ ДОКУМЕНТОВ")
        print("-" * 30)
        merge_output_files(output_folder)

        # Итоговая статистика
        print(f"\n🎉 === ОБРАБОТКА ЗАВЕРШЕНА ===")
        print(f"📊 Статистика:")
        print(f"   • Обработано строк: {processed_rows}")
        print(f"   • Создано файлов: {file_index - 1}")
        print(f"   • Время завершения: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}")
        print("=" * 50)

    except Exception as e:
        print(f"\n💥 КРИТИЧЕСКАЯ ОШИБКА")
        print("-" * 30)
        print(f"❌ {str(e)}")
        messagebox.showerror("Ошибка обработки", f"Произошла ошибка:\n\n{str(e)}")


class TextRedirector:
    """Улучшенный перенаправитель вывода с цветовым выделением"""

    def __init__(self, text_widget):
        self.text_widget = text_widget

        # Настройка тегов для цветового выделения
        text_widget.tag_configure("success", foreground="green")
        text_widget.tag_configure("warning", foreground="orange")
        text_widget.tag_configure("error", foreground="red")
        text_widget.tag_configure("info", foreground="blue")
        text_widget.tag_configure("header", foreground="purple", font=("Arial", 9, "bold"))

    def write(self, text):
        # Определение типа сообщения и применение соответствующего тега
        if "✅" in text or "🎉" in text:
            tag = "success"
        elif "⚠️" in text or "❌" in text:
            tag = "warning" if "⚠️" in text else "error"
        elif "===" in text or "---" in text:
            tag = "header"
        elif "🔍" in text or "ℹ️" in text or "📊" in text:
            tag = "info"
        else:
            tag = None

        self.text_widget.insert(tk.END, text, tag)
        self.text_widget.see(tk.END)

    def flush(self):
        pass


def create_enhanced_log_window(parent, title="📋 Журнал выполнения"):
    """Создает улучшенное окно логов с прогресс-баром"""
    log_window = tk.Toplevel(parent)
    log_window.title(title)
    log_window.geometry("1200x700")
    log_window.resizable(True, True)

    # Верхняя панель с информацией
    top_frame = ttk.Frame(log_window)
    top_frame.pack(fill=tk.X, padx=10, pady=5)

    ttk.Label(top_frame, text="🔄 Обработка документов в процессе...",
              font=("Arial", 12, "bold")).pack(side=tk.LEFT)

    # Прогресс-бар (пока что декоративный)
    progress_frame = ttk.Frame(log_window)
    progress_frame.pack(fill=tk.X, padx=10, pady=5)

    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(progress_frame, mode='indeterminate')
    progress_bar.pack(fill=tk.X)
    progress_bar.start(10)  # Анимированный прогресс-бар

    # Область логов
    log_frame = ttk.LabelFrame(log_window, text="📄 Подробный журнал", padding=5)
    log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    text_log = tk.Text(log_frame, width=100, height=30, font=("Consolas", 9))
    text_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scroll_bar = ttk.Scrollbar(log_frame, command=text_log.yview)
    scroll_bar.pack(side=tk.RIGHT, fill=tk.Y)
    text_log.configure(yscrollcommand=scroll_bar.set)

    # Нижняя панель с кнопкой
    bottom_frame = ttk.Frame(log_window)
    bottom_frame.pack(fill=tk.X, padx=10, pady=5)

    def save_log():
        """Сохранение лога в файл"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")]
        )
        if filename:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(text_log.get(1.0, tk.END))
            messagebox.showinfo("Сохранение", f"Журнал сохранен в:\n{filename}")

    ttk.Button(bottom_frame, text="💾 Сохранить журнал",
               command=save_log).pack(side=tk.LEFT)

    status_label = ttk.Label(bottom_frame, text="Готов к работе...",
                             font=("Arial", 9))
    status_label.pack(side=tk.RIGHT)

    return log_window, text_log, progress_bar, status_label


def main(parent, callback):
    """Главная функция с улучшенным интерфейсом"""
    # Выбор настроек
    selection = choose_files_and_folders(parent, callback)
    if selection is None:
        callback()
        return

    (passports_folder, lab_folder, executive_folder, output_folder,
     excel_file, word_template, double_sided_print, black_and_white) = selection

    # Создание улучшенного окна логов
    log_window, text_log, progress_bar, status_label = create_enhanced_log_window(parent)

    # Перенаправление вывода
    original_stdout = sys.stdout
    original_stderr = sys.stderr
    sys.stdout = TextRedirector(text_log)
    sys.stderr = TextRedirector(text_log)

    def restore_output():
        sys.stdout = original_stdout
        sys.stderr = original_stderr

    def background_job():
        pythoncom.CoInitialize()
        try:
            status_label.config(text="⚙️ Обработка выполняется...")
            run_processing(passports_folder, lab_folder, executive_folder,
                           output_folder, excel_file, word_template,
                           double_sided_print, black_and_white)
        except Exception as e:
            print(f"💥 Необработанная ошибка: {str(e)}")
        finally:
            pythoncom.CoUninitialize()
            restore_output()

            # Остановка прогресс-бара и финальное уведомление
            progress_bar.stop()
            progress_bar.config(mode='determinate', value=100)
            status_label.config(text="✅ Обработка завершена!")

            log_window.after(0, lambda: [
                messagebox.showinfo("🎉 Успешно!",
                                    "Обработка документов завершена!\n\n"
                                    "Проверьте папку вывода для получения результатов.\n"
                                    "Все изображения были автоматически преобразованы в PDF."),
                callback(),
                log_window.destroy()
            ])

    def on_log_window_close():
        restore_output()
        callback()
        log_window.destroy()

    log_window.protocol("WM_DELETE_WINDOW", on_log_window_close)
    threading.Thread(target=background_job, daemon=True).start()


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    root.withdraw()
    root.title("🏢 Система обработки документов")

    # Центрирование главного окна
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (400 // 2)
    y = (root.winfo_screenheight() // 2) - (300 // 2)
    root.geometry(f"400x300+{x}+{y}")


    def on_complete():
        root.quit()


    main(root, on_complete)
    root.mainloop()
