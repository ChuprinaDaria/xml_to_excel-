# xml_to_excel-
Smal code for searching specific info in XML files and creating one excel file with specific info 
import os
import pandas as pd
import xml.etree.ElementTree as ET
from tkinter import Tk, filedialog, messagebox

def parse_xml(file_path):
    """Функція для парсингу одного XML-файлу."""
    # Отримання MRN з назви файлу без префіксу CC599C-
    mrn = os.path.basename(file_path).replace("CC599C-", "").split(".")[0]

    # Парсинг XML
    tree = ET.parse(file_path)
    root = tree.getroot()

    # Отримання даних
    goods_items = []
    for item in root.findall(".//GoodsItem"):
        nr_towaru = item.find("declarationGoodsItemNumber").text if item.find("declarationGoodsItemNumber") is not None else "N/A"
        opis_towaru = item.find(".//descriptionOfGoods").text if item.find(".//descriptionOfGoods") is not None else "N/A"
        nr_listu = item.find(".//TransportDocument/referenceNumber").text if item.find(".//TransportDocument/referenceNumber") is not None else "N/A"

        goods_items.append({
            "MRN": mrn,
            "Nr Towaru": nr_towaru,
            "Opis Towaru": opis_towaru,
            "Nr Listu": nr_listu
        })
    return goods_items

def process_files():
    """Функція для вибору файлів і обробки."""
    # Вибір файлів
    file_paths = filedialog.askopenfilenames(title="Виберіть XML-файли", filetypes=[("XML Files", "*.xml")])
    if not file_paths:
        messagebox.showwarning("Увага", "Файли не вибрано!")
        return

    # Вибір місця для збереження
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if not output_file:
        messagebox.showwarning("Увага", "Шлях для збереження не вибрано!")
        return

    all_data = []
    for file_path in file_paths:
        all_data.extend(parse_xml(file_path))

    # Конвертуємо дані в DataFrame
    df = pd.DataFrame(all_data)

    # Зберігаємо в Excel
    df.to_excel(output_file, index=False)
    messagebox.showinfo("Успіх", f"Дані збережено у файл: {output_file}")

if __name__ == "__main__":
    # Інтерфейс
    root = Tk()
    root.withdraw()  # Приховуємо основне вікно
    process_files()
