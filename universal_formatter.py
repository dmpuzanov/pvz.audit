# -*- coding: utf-8 -*-
"""
UniversalDocFormatter — Строгий форматтер документов по ТЗ
Все параметры взяты ИСКЛЮЧИТЕЛЬНО из предоставленного ТЗ.
"""

import re
import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================================================
# 1. КОНФИГУРАЦИЯ (СТРОГО ПО ТЗ)
# ============================================================
class Config:
    # Поля страницы
    LEFT_MARGIN = Cm(3.0)
    RIGHT_MARGIN = Cm(2.0)
    TOP_MARGIN = Cm(2.0)
    BOTTOM_MARGIN = Cm(2.0)
    
    # Шрифты
    FONT_MAIN = "Georgia"
    FONT_TABLE = "Source Sans Pro"
    
    # Размеры
    SIZE_TEXT = Pt(11)
    SIZE_TABLE = Pt(8)
    SIZE_H1 = Pt(14)
    SIZE_H2 = Pt(13)
    SIZE_H3 = Pt(12)
    SIZE_FOOTER = Pt(10)
    
    # Интервалы и отступы
    LINE_SPACING = 1.0
    FIRST_LINE_INDENT = Cm(0)
    SPACE_BEFORE_H1 = Pt(12)
    SPACE_AFTER_H1 = Pt(12)
    SPACE_BEFORE_H23 = Pt(6)
    SPACE_AFTER_H23 = Pt(6)
    
    # Типографика
    DOUBLE_SPACE_AFTER_PERIOD = True
    LANGUAGE = "ru-RU"
    
    # Таблицы
    TABLE_HEADER_BG = "F2F2F2"
    TABLE_ALTERNATE_ROWS = False

# ============================================================
# 2. ТИПОГРАФИЧЕСКИЙ ДВИЖОК
# ============================================================
class Typography:
    @staticmethod
    def apply_nbsp(text):
        # Используем chr(160) вместо \u00a0 для избежания ошибок в raw-строках
        nbsp = chr(160)
        
        rules = [
            # Инициалы и фамилия
            (r'([А-Я])\. ([А-Я][а-я]+)', r'\1.' + nbsp + r'\2'),
            # Сокращения с точкой
            (r'(г|п|ст|ч|рис|табл|гл|кв|д|корп)\. ', r'\1.' + nbsp + ' '),
            # Нумерация
            (r'(№)\s*', r'\1' + nbsp),
            # Организационно-правовые формы
            (r'(ООО|АО|ИП|ПАО|ЗАО|ФЗ|ГК|АПК|УК|НК|РФ)\s+', r'\1' + nbsp),
            # Даты с "по"
            (r'(\d{1,2}\.\d{1,2}\.\d{4})\s+(по|с)\s+(\d{1,2}\.\d{1,2}\.\d{4})', 
             lambda m: m.group(1) + nbsp + m.group(2) + nbsp + m.group(3)),
            # Суммы с валютой
            (r'(\d[\d\s.,]{1,20})(₽|руб\.|тыс\.|млн\.)', r'\1' + nbsp + r'\2'),
            # Разделитель тысяч
            (r'(\d) (\d{3})', r'\1' + nbsp + r'\2'),
        ]
        
        for pat, repl in rules:
            if callable(repl):
                text = re.sub(pat, repl, text, flags=re.IGNORECASE)
            else:
                text = re.sub(pat, repl, text, flags=re.IGNORECASE)
        return text

    @staticmethod
    def apply_quotes(text):
        return re.sub(r'"([^"]*)"', r'«\1»', text)

    @staticmethod
    def apply_dashes(text):
        # Длинное тире с пробелами (используем chr для безопасной вставки Unicode)
        em_dash = chr(0x2014)  # —
        en_dash = chr(0x2013)  # –
        
        # Длинное тире с пробелами
        text = re.sub(r'\s+-\s+', f' {em_dash} ', text)
        # Диапазон (короткое тире без пробелов)
        text = re.sub(r'(\d+)\s*-\s*(\d+)', rf'\1{en_dash}\2', text)
        return text

    @staticmethod
    def double_space_after_period(text):
        # Двойной пробел после точки, но не после сокращений
        # Используем простой подход без look-behind переменной длины
        # Сначала обрабатываем основные сокращения
        protected = [
            r'т\.е\.', r'т\.к\.', r'и\.т\.д', r'и\.т\.п', 
            r'см\.', r'рис\.', r'табл\.', r'гл\.', 
            r'п\.', r'ст\.', r'д\.', r'кв\.', r'корп\.',
            r'ООО', r'АО', r'ИП', r'ПАО', r'ЗАО', r'№', r'г\.'
        ]
        # Временная замена сокращений
        for i, abbr in enumerate(protected):
            text = re.sub(abbr + r'\s+', f'__ABBR{i}__', text)
        
        # Добавляем двойной пробел после точки
        text = re.sub(r'([.!?])\s+', r'\1  ', text)
        
        # Возвращаем сокращения обратно
        for i, abbr in enumerate(protected):
            text = re.sub(f'__ABBR{i}__', abbr + ' ', text)
        
        return text

    @staticmethod
    def process(text):
        text = Typography.apply_quotes(text)
        text = Typography.apply_dashes(text)
        text = Typography.apply_nbsp(text)
        if Config.DOUBLE_SPACE_AFTER_PERIOD:
            text = Typography.double_space_after_period(text)
        return text

# ============================================================
# 3. XML-УТИЛИТЫ
# ============================================================
class XmlUtils:
    @staticmethod
    def set_cell_shading(cell, hex_color):
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), hex_color)
        shd.set(qn('w:val'), 'clear')
        cell._element.get_or_add_tcPr().append(shd)

    @staticmethod
    def set_run_lang(run, lang="ru-RU"):
        rPr = run._element.get_or_add_rPr()
        lang_el = OxmlElement('w:lang')
        lang_el.set(qn('w:val'), lang)
        lang_el.set(qn('w:eastAsia'), lang)
        rPr.append(lang_el)

    @staticmethod
    def enable_hyphenation(doc):
        try:
            hyph = OxmlElement('w:hyphenation')
            hyph.set(qn('w:auto'), '1')
            doc.settings._element.append(hyph)
        except: 
            pass

    @staticmethod
    def set_keep_with_next(paragraph):
        pPr = paragraph._element.get_or_add_pPr()
        keep = OxmlElement('w:keepNext')
        keep.set(qn('w:val'), '1')
        pPr.append(keep)

    @staticmethod
    def set_keep_together(paragraph):
        pPr = paragraph._element.get_or_add_pPr()
        keep = OxmlElement('w:keepLines')
        keep.set(qn('w:val'), '1')
        pPr.append(keep)

# ============================================================
# 4. ПАРСЕР РАЗМЕТКИ
# ============================================================
class Parser:
    @staticmethod
    def parse(raw_text):
        lines = raw_text.split('\n')
        structure = []
        in_table, t_headers, t_rows = False, [], []
        
        for line in lines:
            line = line.rstrip()
            if not line:
                structure.append(('empty', ''))
                continue
            
            if line.startswith('# '): 
                structure.append(('h1', line[2:].strip()))
            elif line.startswith('## '): 
                structure.append(('h2', line[3:].strip()))
            elif line.startswith('### '): 
                structure.append(('h3', line[4:].strip()))
            elif line.startswith('|') and '|' in line:
                cells = [c.strip() for c in line.split('|')[1:-1]]
                if not in_table: 
                    in_table, t_headers = True, cells
                else: 
                    t_rows.append(cells)
            elif line.startswith('* ') or line.startswith('- '): 
                structure.append(('bullet', line[2:].strip()))
            elif re.match(r'^\d+\.\s', line): 
                structure.append(('numbered', line.strip()))
            elif in_table:
                structure.append(('table', (t_headers, t_rows)))
                in_table, t_headers, t_rows = False, [], []
                structure.append(('paragraph', line))
            else: 
                structure.append(('paragraph', line))
                
        if in_table and t_rows: 
            structure.append(('table', (t_headers, t_rows)))
        return structure

# ============================================================
# 5. БИЛДЕР ДОКУМЕНТА
# ============================================================
class DocBuilder:
    def __init__(self):
        self.doc = Document()
        self._setup_page()
        self._setup_styles()
        XmlUtils.enable_hyphenation(self.doc)
        
    def _setup_page(self):
        sec = self.doc.sections[0]
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        sec.left_margin = Config.LEFT_MARGIN
        sec.right_margin = Config.RIGHT_MARGIN
        sec.top_margin = Config.TOP_MARGIN
        sec.bottom_margin = Config.BOTTOM_MARGIN
        
        sec.different_first_page_header_footer = True
        footer = sec.footer
        footer.is_linked_to_previous = False
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        fld = OxmlElement('w:fldSimple')
        fld.set(qn('w:instr'), ' PAGE ')
        run._element.append(fld)
        run.font.size = Config.SIZE_FOOTER
        run.font.name = Config.FONT_MAIN

    def _setup_styles(self):
        style = self.doc.styles['Normal']
        style.font.name = Config.FONT_MAIN
        style.font.size = Config.SIZE_TEXT
        style.paragraph_format.line_spacing = Config.LINE_SPACING
        style.paragraph_format.first_line_indent = Config.FIRST_LINE_INDENT
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        style.paragraph_format.widow_control = True

    def _add_run(self, paragraph, text, bold=False, italic=False, font_size=None, font_name=None):
        processed = Typography.process(text)
        run = paragraph.add_run(processed)
        run.font.name = font_name or Config.FONT_MAIN
        run.font.size = font_size or Config.SIZE_TEXT
        run.font.bold = bold
        run.font.italic = italic
        XmlUtils.set_run_lang(run)
        return run

    def add_heading(self, text, level=1):
        p = self.doc.add_paragraph()
        size = Config.SIZE_H1 if level==1 else Config.SIZE_H2 if level==2 else Config.SIZE_H3
        self._add_run(p, text, bold=True, font_size=size)
        
        if level == 1:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Config.SPACE_BEFORE_H1
            p.paragraph_format.space_after = Config.SPACE_AFTER_H1
        else:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Config.SPACE_BEFORE_H23
            p.paragraph_format.space_after = Config.SPACE_AFTER_H23
            
        XmlUtils.set_keep_with_next(p)
        return p

    def add_paragraph(self, text):
        p = self.doc.add_paragraph()
        self._add_run(p, text)
        return p

    def add_bullet_list(self, item):
        p = self.doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.5)
        p.paragraph_format.first_line_indent = Cm(-0.5)
        self._add_run(p, "• " + item)

    def add_table(self, headers, rows):
        table = self.doc.add_table(rows=1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.style = 'Table Grid'
        
        XmlUtils.set_keep_together(table.rows[0].cells[0].paragraphs[0])

        for i, h in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.text = ""
            self._add_run(cell.paragraphs[0], h, bold=True, font_size=Config.SIZE_TABLE, font_name=Config.FONT_TABLE)
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            XmlUtils.set_cell_shading(cell, Config.TABLE_HEADER_BG)

        for r_idx, r_data in enumerate(rows):
            row_cells = table.add_row().cells
            for j, val in enumerate(r_data):
                cell = row_cells[j]
                cell.text = ""
                is_num = bool(re.match(r'^[\d\s.,₽-]+$', str(val)))
                self._add_run(cell.paragraphs[0], str(val), font_size=Config.SIZE_TABLE, font_name=Config.FONT_TABLE)
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT if is_num else WD_ALIGN_PARAGRAPH.LEFT
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            
            if Config.TABLE_ALTERNATE_ROWS and r_idx % 2 == 0:
                for c in row_cells:
                    XmlUtils.set_cell_shading(c, "F9F9F9")

        tbl = table._element
        for row in table.rows:
            for cell in row.cells:
                tc = cell._element
                tcPr = tc.get_or_add_tcPr()
                borders = OxmlElement('w:tcBorders')
                for b_name in ['top', 'bottom']:
                    b = OxmlElement('w:' + b_name)
                    b.set(qn('w:val'), 'single')
                    b.set(qn('w:sz'), '4')
                    b.set(qn('w:color'), '000000')
                    borders.append(b)
                tcPr.append(borders)
        return table

    def add_title_page(self, title, author, city, date):
        for _ in range(10): 
            self.doc.add_paragraph()
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._add_run(p, title, bold=True, font_size=Pt(18))
        
        p2 = self.doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        self._add_run(p2, author + "\n" + city + "\n" + date, italic=True, font_size=Pt(14))
        
        self.doc.add_page_break()

    def validate(self):
        errors = []
        for para in self.doc.paragraphs:
            if para.paragraph_format.line_spacing is not None and para.paragraph_format.line_spacing != Config.LINE_SPACING:
                errors.append("Неверный интервал: " + para.text[:40])
            if para.paragraph_format.first_line_indent != Config.FIRST_LINE_INDENT:
                errors.append("Неверный отступ: " + para.text[:40])
        for table in self.doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for run in cell.paragraphs[0].runs:
                        if run.font.size != Config.SIZE_TABLE and run.font.name != Config.FONT_TABLE:
                            errors.append("Шрифт таблицы не совпадает с ТЗ")
        return errors

    def save(self, path):
        self.doc.save(path)

# ============================================================
# 6. GUI
# ============================================================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("UniversalDocFormatter (ТЗ v1.0)")
        self.root.geometry("1100x800")
        
        tk.Label(root, text="ВВЕДИТЕ ТЕКСТ", font=("Arial", 12, "bold"), fg="#2c3e50").pack(pady=5)
        self.text_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, font=("Consolas", 11))
        self.text_area.pack(padx=15, pady=10, fill=tk.BOTH, expand=True)
        
        ctrl_frame = tk.Frame(root)
        ctrl_frame.pack(pady=5)
        self.var_title = tk.BooleanVar()
        tk.Checkbutton(ctrl_frame, text="Титульная страница", variable=self.var_title).pack(side=tk.LEFT, padx=10)
        self.var_alt_rows = tk.BooleanVar()
        tk.Checkbutton(ctrl_frame, text="Чередование строк таблиц", variable=self.var_alt_rows).pack(side=tk.LEFT, padx=10)
        
        tk.Label(root, text="# Заголовок 1 | ## Заголовок 2 | ### Заголовок 3 | * список | | таблица |", fg="gray").pack()
        tk.Button(root, text="ОТФОРМАТИРОВАТЬ И СОХРАНИТЬ", command=self.process, bg="#2980b9", fg="white", font=("Arial", 11, "bold")).pack(pady=10)

    def process(self):
        raw = self.text_area.get("1.0", tk.END).strip()
        if not raw:
            messagebox.showwarning("Внимание", "Текст пустой")
            return
            
        try:
            Config.TABLE_ALTERNATE_ROWS = self.var_alt_rows.get()
            builder = DocBuilder()
            
            if self.var_title.get():
                builder.add_title_page("Название документа", "Автор", "г. Москва", "01.01.2026")
                
            structure = Parser.parse(raw)
            for typ, data in structure:
                if typ == 'empty': 
                    builder.doc.add_paragraph()
                elif typ in ('h1','h2','h3'): 
                    builder.add_heading(data, level={'h1':1,'h2':2,'h3':3}[typ])
                elif typ == 'paragraph': 
                    builder.add_paragraph(data)
                elif typ == 'bullet': 
                    builder.add_bullet_list(data)
                elif typ == 'table': 
                    builder.add_table(data[0], data[1])
                
            errs = builder.validate()
            if errs: 
                messagebox.showwarning("Валидация", "\n".join(errs[:5]))
            
            path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")])
            if path:
                builder.save(path)
                messagebox.showinfo("Готово", "Сохранено: " + path)
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
