import xml.etree.ElementTree as ET
from html.entities import name2codepoint
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
from pathlib import Path
import html
import re
import urllib.request
from urllib.error import URLError, HTTPError
import pyperclip


class XMLToXLSXConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("YML прайс в XLSX")

        base_dir = Path(__file__).resolve().parent
        my_icon = base_dir / "app.ico"
        if my_icon.exists():
            try:
                self.root.iconbitmap(str(my_icon))
            except Exception:
                pass

        self.root.geometry("720x620")

        self.xml_file_path = tk.StringVar()
        self.xml_url = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.output_filename = tk.StringVar(value="products.xlsx")
        self.fix_esc_sequences = tk.BooleanVar(value=True)
        self.source_type = tk.StringVar(value="url")
        self.duplicate_handling = tk.StringVar(value="separate")
        self.encoding = tk.StringVar(value="auto")

        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        title_label = ttk.Label(main_frame, text="YML прайс конвертер", font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=4, pady=(0, 15))

        source_frame = ttk.LabelFrame(main_frame, text="Источник данных", padding="5")
        source_frame.grid(row=1, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
        source_frame.columnconfigure(1, weight=1)

        ttk.Radiobutton(
            source_frame,
            text="Локальный файл",
            variable=self.source_type,
            value="local",
            command=self.toggle_source_type,
        ).grid(row=0, column=0, sticky=tk.W)

        ttk.Radiobutton(
            source_frame,
            text="URL (ссылка)",
            variable=self.source_type,
            value="url",
            command=self.toggle_source_type,
        ).grid(row=1, column=0, sticky=tk.W)

        self.local_frame = ttk.Frame(source_frame)
        self.local_frame.grid(row=0, column=1, columnspan=3, sticky=(tk.W, tk.E), pady=2)
        self.local_frame.columnconfigure(0, weight=1)

        ttk.Entry(self.local_frame, textvariable=self.xml_file_path, state="readonly").grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5)
        )
        ttk.Button(self.local_frame, text="Обзор", command=self.browse_xml_file).grid(row=0, column=1)

        self.url_frame = ttk.Frame(source_frame)
        self.url_frame.grid(row=1, column=1, columnspan=3, sticky=(tk.W, tk.E), pady=2)
        self.url_frame.columnconfigure(0, weight=1)

        ttk.Entry(self.url_frame, textvariable=self.xml_url).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5)
        )

        url_buttons_frame = ttk.Frame(self.url_frame)
        url_buttons_frame.grid(row=0, column=1)

        ttk.Button(url_buttons_frame, text="Ctrl+V", command=self.paste_url, width=7).grid(
            row=0, column=0, padx=(0, 2)
        )
        ttk.Button(url_buttons_frame, text="Тест", command=self.test_url, width=7).grid(row=0, column=1)

        ttk.Label(main_frame, text="Папка для сохранения:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_folder, state="readonly").grid(
            row=2, column=1, sticky=(tk.W, tk.E), pady=5, padx=(0, 5), columnspan=2
        )
        ttk.Button(main_frame, text="Обзор", command=self.browse_output_folder).grid(row=2, column=3, pady=5)

        ttk.Label(main_frame, text="Имя файла:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(main_frame, textvariable=self.output_filename).grid(
            row=3, column=1, sticky=(tk.W, tk.E), pady=5, columnspan=3
        )

        ttk.Label(main_frame, text="Кодировка XML:").grid(row=4, column=0, sticky=tk.W, pady=5)

        encoding_frame = ttk.Frame(main_frame)
        encoding_frame.grid(row=4, column=1, columnspan=3, sticky=tk.W, pady=5)

        ttk.Radiobutton(encoding_frame, text="Автоопределение", variable=self.encoding, value="auto").pack(
            side=tk.LEFT, padx=5
        )
        ttk.Radiobutton(encoding_frame, text="UTF-8", variable=self.encoding, value="utf-8").pack(
            side=tk.LEFT, padx=5
        )
        ttk.Radiobutton(
            encoding_frame,
            text="Windows-1251 (CP1251)",
            variable=self.encoding,
            value="cp1251",
        ).pack(side=tk.LEFT, padx=5)

        duplicate_frame = ttk.LabelFrame(main_frame, text="Обработка повторяющихся полей", padding="5")
        duplicate_frame.grid(row=5, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=10)

        ttk.Radiobutton(
            duplicate_frame,
            text="Объединять через ';' (значения в одном столбце)",
            variable=self.duplicate_handling,
            value="merge",
        ).pack(side=tk.LEFT, padx=10)

        ttk.Radiobutton(
            duplicate_frame,
            text="Создавать отдельные столбцы (picture_1, picture_2)",
            variable=self.duplicate_handling,
            value="separate",
        ).pack(side=tk.LEFT, padx=10)

        esc_checkbox = ttk.Checkbutton(
            main_frame,
            text='Исправлять ESC-последовательности (&amp;, &nbsp;, &quot; и т.д.)',
            variable=self.fix_esc_sequences,
        )
        esc_checkbox.grid(row=6, column=0, columnspan=4, sticky=tk.W, pady=10)

        help_label = ttk.Label(
            main_frame,
            text='ВКЛ: &amp; → &, &nbsp; → пробел, &quot; → ", HTML-сущности декодируются',
            font=("Arial", 8),
            foreground="gray",
        )
        help_label.grid(row=7, column=0, columnspan=4, sticky=tk.W, pady=(0, 10))

        self.progress = ttk.Progressbar(main_frame, mode="indeterminate")
        self.progress.grid(row=8, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=20)

        self.log_text = tk.Text(main_frame, height=10, width=70)
        self.log_text.grid(row=9, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)

        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=9, column=4, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=scrollbar.set)

        self.convert_button = ttk.Button(main_frame, text="Конвертировать", command=self.start_conversion)
        self.convert_button.grid(row=10, column=0, columnspan=4, pady=10)

        main_frame.rowconfigure(9, weight=1)
        for i in range(4):
            main_frame.columnconfigure(i, weight=1)

        self.toggle_source_type()

    def toggle_source_type(self):
        if self.source_type.get() == "local":
            self.local_frame.grid()
            self.url_frame.grid_remove()
        else:
            self.local_frame.grid_remove()
            self.url_frame.grid()

    def paste_url(self):
        try:
            clipboard_text = pyperclip.paste()
            if clipboard_text:
                self.xml_url.set(clipboard_text.strip())
                self.log("URL вставлен из буфера обмена")
            else:
                self.log("Буфер обмена пуст")
        except Exception as e:
            self.log(f"Ошибка при вставке из буфера обмена: {e}")
            messagebox.showerror("Ошибка", f"Не удалось вставить из буфера обмена:\n{e}")

    def browse_xml_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите XML файл",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
        )
        if filename:
            self.xml_file_path.set(filename)
            self.log(f"Выбран XML файл: {filename}")

    def browse_output_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку для сохранения")
        if folder:
            self.output_folder.set(folder)
            self.log(f"Выбрана папка для сохранения: {folder}")

    def test_url(self):
        url = self.xml_url.get().strip()
        if not url:
            messagebox.showwarning("Предупреждение", "Введите URL для проверки")
            return

        if not url.startswith(("http://", "https://")):
            url = "http://" + url

        self.log(f"Проверка URL: {url}")

        try:
            response = urllib.request.urlopen(url, timeout=10)
            if response.getcode() == 200:
                self.log("URL доступен!")
                messagebox.showinfo("Успех", "URL доступен и готов к загрузке")
            else:
                self.log(f"Сервер вернул код: {response.getcode()}")
        except Exception as e:
            self.log(f"Ошибка при проверке URL: {e}")
            messagebox.showerror("Ошибка", f"URL недоступен:\n{e}")

    def download_xml_from_url(self, url, temp_file):
        try:
            self.log(f"Загрузка XML из: {url}")
            headers = {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/91.0.4472.124 Safari/537.36"
                )
            }
            req = urllib.request.Request(url, headers=headers)

            with urllib.request.urlopen(req, timeout=60) as response:
                content = response.read()

            if b"<?xml" not in content[:200].lower() and b"<yml_catalog" not in content[:2000].lower():
                self.log("Внимание: загруженный файл может не быть XML")

            with open(temp_file, "wb") as f:
                f.write(content)

            self.log(f"XML загружен успешно, размер: {len(content)} байт")
            return True
        except HTTPError as e:
            self.log(f"HTTP ошибка: {e.code} - {e.reason}")
            raise Exception(f"HTTP ошибка: {e.code} - {e.reason}")
        except URLError as e:
            self.log(f"URL ошибка: {e.reason}")
            raise Exception(f"URL ошибка: {e.reason}")
        except Exception as e:
            self.log(f"Ошибка загрузки: {e}")
            raise Exception(f"Ошибка загрузки: {e}")

    def log(self, message):
        self.log_text.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def detect_encoding(self, xml_file):
        self.log("Определение кодировки XML файла...")
        try:
            with open(xml_file, "rb") as f:
                raw_data = f.read(1000)

            result = re.search(br'encoding=["\']([^"\']+)["\']', raw_data)
            if result:
                detected_encoding = result.group(1).decode("ascii", errors="ignore").lower()
                self.log(f"Кодировка указана в XML: {detected_encoding}")
                return detected_encoding

            try:
                raw_data.decode("utf-8")
                self.log("Определена кодировка: UTF-8")
                return "utf-8"
            except UnicodeDecodeError:
                pass

            try:
                raw_data.decode("cp1251")
                self.log("Определена кодировка: Windows-1251")
                return "cp1251"
            except UnicodeDecodeError:
                pass

            for encoding in ["utf-8-sig", "iso-8859-1", "windows-1252"]:
                try:
                    raw_data.decode(encoding)
                    self.log(f"Определена кодировка: {encoding}")
                    return encoding
                except UnicodeDecodeError:
                    continue

            self.log("Не удалось определить кодировку, используем UTF-8")
            return "utf-8"
        except Exception as e:
            self.log(f"Ошибка при определении кодировки: {e}")
            return "utf-8"

    def read_xml_with_encoding(self, xml_file):
        encoding_setting = self.encoding.get()

        if encoding_setting == "auto":
            actual_encoding = self.detect_encoding(xml_file)
        else:
            actual_encoding = encoding_setting
            self.log(f"Используется принудительная кодировка: {actual_encoding}")

        try:
            with open(xml_file, "r", encoding=actual_encoding) as file:
                return file.read()
        except UnicodeDecodeError as e:
            self.log(f"Ошибка декодирования с {actual_encoding}: {e}")
            self.log("Пробуем альтернативные кодировки...")

            alternative_encodings = ["utf-8", "cp1251", "utf-8-sig", "iso-8859-1", "windows-1252"]
            for alt_encoding in alternative_encodings:
                if alt_encoding == actual_encoding:
                    continue
                try:
                    with open(xml_file, "r", encoding=alt_encoding) as file:
                        xml_content = file.read()
                    self.log(f"Успешно прочитан файл с кодировкой: {alt_encoding}")
                    return xml_content
                except UnicodeDecodeError:
                    continue

            try:
                with open(xml_file, "rb") as file:
                    binary_content = file.read()
                xml_content = binary_content.decode("utf-8", errors="ignore")
                self.log("Файл прочитан с игнорированием ошибок декодирования")
                return xml_content
            except Exception as e2:
                raise Exception(f"Не удалось прочитать файл ни в одной кодировке: {e2}")

    def start_conversion(self):
        if self.source_type.get() == "local" and not self.xml_file_path.get():
            messagebox.showerror("Ошибка", "Выберите XML файл")
            return

        if self.source_type.get() == "url" and not self.xml_url.get().strip():
            messagebox.showerror("Ошибка", "Введите URL")
            return

        if not self.output_folder.get():
            messagebox.showerror("Ошибка", "Выберите папку для сохранения")
            return

        if not self.output_filename.get():
            messagebox.showerror("Ошибка", "Введите имя выходного файла")
            return

        self.convert_button.config(state="disabled")
        self.progress.start()

        thread = threading.Thread(target=self.convert_xml_to_xlsx)
        thread.daemon = True
        thread.start()

    def convert_xml_to_xlsx(self):
        temp_file = None
        try:
            output_path = os.path.join(self.output_folder.get(), self.output_filename.get())

            self.log("=== НАЧАЛО КОНВЕРТАЦИИ ===")
            self.log(f"Источник: {'URL' if self.source_type.get() == 'url' else 'локальный файл'}")
            self.log(f"Кодировка: {self.encoding.get()}")
            self.log(
                f"Исправление ESC-последовательностей: {'ВКЛ' if self.fix_esc_sequences.get() else 'ВЫКЛ'}"
            )
            self.log(
                "Режим обработки дубликатов: "
                + ("объединение" if self.duplicate_handling.get() == "merge" else "отдельные столбцы")
            )

            if self.source_type.get() == "url":
                url = self.xml_url.get().strip()
                if not url.startswith(("http://", "https://")):
                    url = "http://" + url

                temp_file = os.path.join(self.output_folder.get(), "temp_xml_download.xml")
                self.download_xml_from_url(url, temp_file)
                xml_source = temp_file
            else:
                xml_source = self.xml_file_path.get()

            if self.duplicate_handling.get() == "merge":
                self.parse_xml_to_xlsx_with_categories_merge(xml_source, output_path)
            else:
                self.parse_xml_to_xlsx_with_categories_separate(xml_source, output_path)

            self.log("=== КОНВЕРТАЦИЯ ЗАВЕРШЕНА ===")
            messagebox.showinfo("Успех", f"Файл успешно создан:\n{output_path}")
        except Exception as e:
            self.log(f"Ошибка: {e}")
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{e}")
        finally:
            if temp_file and os.path.exists(temp_file):
                try:
                    os.remove(temp_file)
                except Exception:
                    pass
            self.root.after(0, self.conversion_finished)

    def conversion_finished(self):
        self.convert_button.config(state="normal")
        self.progress.stop()

    def sanitize_column_name(self, name):
        if not name:
            return ""
        clean_name = re.sub(r"\s+", "_", str(name).strip())
        clean_name = re.sub(r"[^\w.-]", "_", clean_name, flags=re.UNICODE)
        clean_name = re.sub(r"_+", "_", clean_name).strip("_")
        return clean_name or "param"

    def preprocess_xml_content(self, xml_content):
        if not xml_content:
            return xml_content

        text = xml_content.replace("\ufeff", "")

        if not self.fix_esc_sequences.get():
            return text

        text = text.replace("&nbsp;", "&#160;").replace("&Nbsp;", "&#160;")
        entity_pattern = re.compile(r"&([A-Za-z][A-Za-z0-9]+);")

        def replace_named_entity(match):
            entity_name = match.group(1)
            if entity_name in {"amp", "lt", "gt", "quot", "apos"}:
                return match.group(0)

            codepoint = name2codepoint.get(entity_name)
            if codepoint is None:
                codepoint = name2codepoint.get(entity_name.lower())

            if codepoint is not None:
                return chr(codepoint)

            return match.group(0)

        return entity_pattern.sub(replace_named_entity, text)

    def parse_xml_root(self, xml_file):
        xml_content = self.read_xml_with_encoding(xml_file)
        xml_content = self.preprocess_xml_content(xml_content)

        try:
            return ET.fromstring(xml_content)
        except ET.ParseError as e:
            preview = ""
            if hasattr(e, "position") and isinstance(e.position, tuple) and len(e.position) >= 2:
                line_no, col_no = e.position
                lines = xml_content.splitlines()
                if 1 <= line_no <= len(lines):
                    line = lines[line_no - 1]
                    preview = line[max(0, col_no - 120): col_no + 120]

            if preview:
                preview = preview.replace("\n", " ").replace("\r", " ")
                self.log(f"Фрагмент возле ошибки XML: {preview}")

            raise Exception(f"Ошибка парсинга XML: {e}")

    def process_text(self, text):
        if text is None:
            return ""

        text = str(text)

        if not self.fix_esc_sequences.get():
            return text

        try:
            text = html.unescape(text)
            text = text.replace("\xa0", " ")
            return text
        except Exception as e:
            self.log(f"Ошибка при обработке текста: {e}")
            return text

    def get_element_raw_text(self, element):
        if element is None:
            return ""
        return "".join(element.itertext())

    def autosize_worksheet_columns(self, writer):
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        cell_value = "" if cell.value is None else str(cell.value)
                        if len(cell_value) > max_length:
                            max_length = len(cell_value)
                    except Exception:
                        pass
                worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)

    def parse_xml_to_xlsx_with_categories_merge(self, xml_file, xlsx_file):
        self.log("Чтение XML файла (режим объединения)...")
        root = self.parse_xml_root(xml_file)

        categories_data = []
        for category in root.findall('.//categories/category'):
            raw_text = self.get_element_raw_text(category)
            categories_data.append({
                'id': category.get('id', ''),
                'name': self.process_text(raw_text) if raw_text else self.process_text(category.text or ''),
            })

        self.log(f"Найдено категорий: {len(categories_data)}")

        all_simple_fields = set()
        all_param_names = set()
        offers_data = []

        self.log("Анализ структуры товаров...")
        offers = root.findall('.//offer')
        total_offers = len(offers)

        for i, offer in enumerate(offers):
            if offer.get('id') is not None:
                all_simple_fields.add('id')
            if offer.get('available') is not None:
                all_simple_fields.add('available')

            for child in offer:
                if child.tag != 'param':
                    all_simple_fields.add(child.tag)

            for param in offer.findall('param'):
                param_name = param.get('name')
                if param_name:
                    all_param_names.add(self.sanitize_column_name(param_name))

            if i % 100 == 0 and total_offers > 100:
                self.log(f"Проанализировано {i}/{total_offers} товаров...")

        self.log(f"Найдено полей: {len(all_simple_fields)} простых + {len(all_param_names)} параметров")
        self.log("Обработка данных товаров...")

        for i, offer in enumerate(offers):
            offer_data = {
                'id': self.process_text(offer.get('id', '')),
                'available': self.process_text(offer.get('available', '')),
            }

            for field in all_simple_fields:
                if field not in ['id', 'available']:
                    elements = offer.findall(field)
                    values = []
                    for element in elements:
                        raw_text = self.get_element_raw_text(element)
                        if raw_text:
                            values.append(self.process_text(raw_text))
                    offer_data[field] = '; '.join(v for v in values if v)

            for param_name in all_param_names:
                offer_data[param_name] = ''

            param_values = {}
            for param in offer.findall('param'):
                param_name = param.get('name')
                if param_name:
                    clean_param_name = self.sanitize_column_name(param_name)
                    raw_text = self.get_element_raw_text(param)
                    value = self.process_text(raw_text) if raw_text else self.process_text(param.text or '')
                    param_values.setdefault(clean_param_name, [])
                    if value:
                        param_values[clean_param_name].append(value)

            for clean_param_name, values in param_values.items():
                offer_data[clean_param_name] = '; '.join(values)

            offers_data.append(offer_data)

            if i % 100 == 0 and total_offers > 100:
                self.log(f"Обработано {i}/{total_offers} товаров...")

        self.log("Создание Excel файла...")
        df_offers = pd.DataFrame(offers_data)
        df_categories = pd.DataFrame(categories_data)

        base_columns = ['id', 'available'] + sorted([f for f in all_simple_fields if f not in ['id', 'available']])
        param_columns = sorted(all_param_names)
        ordered_columns = base_columns + param_columns

        for col in ordered_columns:
            if col not in df_offers.columns:
                df_offers[col] = ''

        df_offers = df_offers[ordered_columns]

        with pd.ExcelWriter(xlsx_file, engine='openpyxl') as writer:
            df_offers.to_excel(writer, sheet_name='Товары', index=False)
            df_categories.to_excel(writer, sheet_name='Категории', index=False)
            self.autosize_worksheet_columns(writer)

        self.log(f"Создан файл {xlsx_file} с {len(offers_data)} товарами и {len(categories_data)} категориями")

    def parse_xml_to_xlsx_with_categories_separate(self, xml_file, xlsx_file):
        self.log("Чтение XML файла (режим отдельных столбцов)...")
        root = self.parse_xml_root(xml_file)

        categories_data = []
        for category in root.findall('.//categories/category'):
            raw_text = self.get_element_raw_text(category)
            categories_data.append({
                'id': category.get('id', ''),
                'name': self.process_text(raw_text) if raw_text else self.process_text(category.text or ''),
            })

        self.log(f"Найдено категорий: {len(categories_data)}")
        offers_data = []
        self.log("Анализ структуры товаров...")

        offers = root.findall('.//offer')
        total_offers = len(offers)
        max_counts = {}

        for i, offer in enumerate(offers):
            tag_counts = {}

            for child in offer:
                if child.tag != 'param':
                    tag_counts[child.tag] = tag_counts.get(child.tag, 0) + 1

            param_counts = {}
            for param in offer.findall('param'):
                param_name = param.get('name')
                if param_name:
                    clean_param_name = self.sanitize_column_name(param_name)
                    param_counts[clean_param_name] = param_counts.get(clean_param_name, 0) + 1

            for tag, count in tag_counts.items():
                if tag not in max_counts or count > max_counts[tag]:
                    max_counts[tag] = count

            for param_name, count in param_counts.items():
                if param_name not in max_counts or count > max_counts[param_name]:
                    max_counts[param_name] = count

            if i % 100 == 0 and total_offers > 100:
                self.log(f"Проанализировано {i}/{total_offers} товаров...")

        self.log(f"Максимальное количество повторений тегов: {max_counts}")
        self.log("Обработка данных товаров...")

        for i, offer in enumerate(offers):
            offer_data = {
                'id': self.process_text(offer.get('id', '')),
                'available': self.process_text(offer.get('available', '')),
            }
            tag_values = {}

            for child in offer:
                if child.tag != 'param':
                    tag_values.setdefault(child.tag, [])
                    raw_text = self.get_element_raw_text(child)
                    tag_values[child.tag].append(self.process_text(raw_text if raw_text else (child.text or '')))

            for param in offer.findall('param'):
                param_name = param.get('name')
                if param_name:
                    clean_param_name = self.sanitize_column_name(param_name)
                    tag_values.setdefault(clean_param_name, [])
                    raw_text = self.get_element_raw_text(param)
                    tag_values[clean_param_name].append(self.process_text(raw_text if raw_text else (param.text or '')))

            for tag, max_count in max_counts.items():
                values = tag_values.get(tag, [])
                for j in range(max_count):
                    column_name = f"{tag}_{j + 1}" if max_count > 1 else tag
                    offer_data[column_name] = values[j] if j < len(values) else ''

            offers_data.append(offer_data)

            if i % 100 == 0 and total_offers > 100:
                self.log(f"Обработано {i}/{total_offers} товаров...")

        self.log("Создание Excel файла...")
        ordered_columns = ['id', 'available']

        base_tags = sorted([tag for tag in max_counts.keys() if tag not in ['id', 'available']])
        for tag in base_tags:
            max_count = max_counts[tag]
            if max_count == 1:
                ordered_columns.append(tag)
            else:
                for j in range(max_count):
                    ordered_columns.append(f"{tag}_{j + 1}")

        df_offers = pd.DataFrame(offers_data)
        df_categories = pd.DataFrame(categories_data)

        for col in ordered_columns:
            if col not in df_offers.columns:
                df_offers[col] = ''

        df_offers = df_offers[ordered_columns]

        with pd.ExcelWriter(xlsx_file, engine='openpyxl') as writer:
            df_offers.to_excel(writer, sheet_name='Товары', index=False)
            df_categories.to_excel(writer, sheet_name='Категории', index=False)
            self.autosize_worksheet_columns(writer)

        self.log(f"Создан файл {xlsx_file} с {len(offers_data)} товарами и {len(categories_data)} категориями")
        self.log(f"Создано колонок: {len(ordered_columns)}")


def main():
    try:
        import pyperclip  # noqa: F401
    except ImportError:
        print("Установите библиотеку pyperclip для работы с буфером обмена:")
        print("pip install pyperclip")
        return

    root = tk.Tk()
    app = XMLToXLSXConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()