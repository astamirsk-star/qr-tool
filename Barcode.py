import customtkinter as ctk
from tkinter import filedialog
import os
import re
import glob
import pandas as pd
import fitz
import threading

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class BarcodeApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Поиск штрихкодов в PDF")
        self.geometry("700x450")

        header = ctk.CTkLabel(self, text="Поиск штрихкодов", font=ctk.CTkFont(size=20, weight="bold"))
        header.pack(pady=(20, 10))

        # Поля для выбора файлов
        self.entry_excel = self.create_file_row("1. Выбрать Excel:", self.pick_excel)
        self.entry_pdf = self.create_file_row("2. Папка с PDF:", self.pick_pdf_in)
        self.entry_out = self.create_file_row("3. Куда сохранить:", self.pick_pdf_out)

        # Кнопка запуска и статус
        self.btn_start = ctk.CTkButton(self, text="Начать поиск", font=ctk.CTkFont(size=15), height=40, command=self.start_process)
        self.btn_start.pack(pady=20)

        self.lbl_status = ctk.CTkLabel(self, text="Готов к работе", text_color="#ffcc00")
        self.lbl_status.pack()

        self.textbox = ctk.CTkTextbox(self, width=650, height=150)
        self.textbox.pack(padx=20, pady=10, fill="both", expand=True)

    def create_file_row(self, label_text, command):
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="x", padx=20, pady=5)
        btn = ctk.CTkButton(frame, text=label_text, width=140, command=command)
        btn.pack(side="left", padx=(0, 10))
        entry = ctk.CTkEntry(frame, width=480)
        entry.pack(side="left", fill="x", expand=True)
        return entry

    def pick_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if path: self.entry_excel.delete(0, 'end'); self.entry_excel.insert(0, path)

    def pick_pdf_in(self):
        path = filedialog.askdirectory()
        if path: self.entry_pdf.delete(0, 'end'); self.entry_pdf.insert(0, path)

    def pick_pdf_out(self):
        path = filedialog.askdirectory()
        if path: self.entry_out.delete(0, 'end'); self.entry_out.insert(0, path)

    def log(self, text):
        self.after(0, lambda: self.textbox.insert("end", text + "\n"))
        self.after(0, lambda: self.textbox.see("end"))

    def set_status(self, text):
        self.after(0, lambda: self.lbl_status.configure(text=text))

    def start_process(self):
        ex_path = self.entry_excel.get()
        pdf_path = self.entry_pdf.get()
        out_path = self.entry_out.get()

        if not ex_path or not pdf_path or not out_path:
            self.set_status("Ошибка: Заполните все пути!")
            return

        self.btn_start.configure(state="disabled")
        self.textbox.delete("1.0", "end")
        threading.Thread(target=self.run_search, args=(ex_path, pdf_path, out_path), daemon=True).start()

    def run_search(self, ex_path, pdf_path, out_path):
        self.set_status("Чтение Excel...")
        try:
            # Читаем все листы и собираем штрихкоды
            dfs = pd.read_excel(ex_path, sheet_name=None, header=None)
            barcodes = set()
            for sheet, df in dfs.items():
                for val in df.values.flatten():
                    if pd.notna(val):
                        b = str(val).strip()
                        if b.endswith('.0'): b = b[:-2] # Убираем лишние нули от чисел
                        if b: barcodes.add(b)

            self.log(f"Найдено уникальных штрихкодов в Excel: {len(barcodes)}")

            pdf_files = glob.glob(os.path.join(pdf_path, '*.pdf'))
            total_found = 0

            for pdf_file in pdf_files:
                filename = os.path.basename(pdf_file)
                self.set_status(f"Проверяю: {filename}")
                
                doc = fitz.open(pdf_file)
                new_doc = fitz.open()
                matched_in_file = 0

                for page_num in range(len(doc)):
                    page = doc[page_num]
                    text = re.sub(r'\s+', '', page.get_text()) # Убираем пробелы для надежности
                    
                    for code in barcodes:
                        if code in text:
                            new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
                            matched_in_file += 1
                            total_found += 1
                            break # Переходим к следующей странице
                
                if matched_in_file > 0:
                    out_name = os.path.join(out_path, f"Found_{filename}")
                    new_doc.save(out_name)
                    self.log(f"Файл {filename}: найдено {matched_in_file} стр.")
                
                new_doc.close()
                doc.close()

            self.set_status("Готово!")
            self.log(f"\nПоиск завершен. Всего найдено страниц: {total_found}")

        except Exception as e:
            self.log(f"Ошибка: {e}")
            self.set_status("Произошла ошибка")
        finally:
            self.after(0, lambda: self.btn_start.configure(state="normal"))

if __name__ == "__main__":
    app = BarcodeApp()
    app.mainloop()
