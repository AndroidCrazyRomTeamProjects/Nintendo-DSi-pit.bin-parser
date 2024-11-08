import argparse
import sqlite3
import struct
import datetime
import openpyxl
import pypyodbc
import tkinter as tk
from tkinter import filedialog, messagebox

# Define constants for sticker types
STICKER_TYPES = {
    1: "⭐",
    2: "🍀",
    3: "❤️"
}

def parse_date(seconds_since_2000):
    base_date = datetime.datetime(2000, 1, 1)
    date = base_date + datetime.timedelta(seconds=seconds_since_2000)
    return date.strftime("%Y-%m-%d %H:%M:%S")

def parse_entry(entry_data):
    timestamp, _, flags = struct.unpack("<I8sI", entry_data)
    date_str = parse_date(timestamp)
    photo_number = (flags >> 11) & 0x7F
    sticker_code = (flags >> 18) & 0x3
    sticker = STICKER_TYPES.get(sticker_code, "")
    return date_str, photo_number, sticker

def parse_pit_file(file_path):
    entries = []
    with open(file_path, "rb") as f:
        f.seek(0x08)
        num_entries = struct.unpack("<H", f.read(2))[0]
        f.seek(0x18)
        for _ in range(num_entries):
            entry_data = f.read(16)
            if len(entry_data) < 16:
                break
            date_str, photo_number, sticker = parse_entry(entry_data)
            if date_str == "2000-01-01 00:00:00":
                break
            entries.append((date_str, photo_number, sticker))
    return entries

def export_to_sqlite(entries, output_name="photos.db"):
    conn = sqlite3.connect(output_name)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS photos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date DATETIME,
            photo_number INTEGER,
            sticker TEXT
        )
    ''')
    cursor.executemany('INSERT INTO photos (date, photo_number, sticker) VALUES (?, ?, ?)', entries)
    conn.commit()
    conn.close()
    print(f"Data exported to SQLite database {output_name}")

def export_to_excel(entries, output_name="photos.xlsx"):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(["Date", "Photo Number", "Sticker"])

    for entry in entries:
        
        date_value = datetime.datetime.strptime(entry[0], "%Y-%m-%d %H:%M:%S")
        sheet.append([date_value, entry[1], entry[2]])

    date_column = 1
    for row in sheet.iter_rows(min_row=2, min_col=date_column, max_col=date_column):
        for cell in row:
            if isinstance(cell.value, datetime.datetime):
                cell.number_format = 'YYYY-MM-DD HH:MM:SS'

    workbook.save(output_name)
    print(f"Data exported to Excel file {output_name}")

def export_to_access(entries, output_name="photos.accdb"):
    connection_string = f"Driver={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={output_name};"
    conn = pypyodbc.connect(connection_string)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS photos (
            ID AUTOINCREMENT PRIMARY KEY,
            Date DATETIME,
            PhotoNumber INTEGER,
            Sticker TEXT
        )
    ''')
    cursor.executemany('INSERT INTO photos (Date, PhotoNumber, Sticker) VALUES (?, ?, ?)', entries)
    conn.commit()
    conn.close()
    print(f"Data exported to Access database {output_name}")

    
def handle_db_click(file_path, root):
    root.destroy()
    entries = parse_pit_file(file_path)
    export_to_sqlite(entries)
def handle_xlsx_click(file_path, root):
    root.destroy()
    entries = parse_pit_file(file_path)
    export_to_excel(entries)
def handle_accdb_click(file_path, root):
    root.destroy()
    entries = parse_pit_file(file_path)
    export_to_access(entries)


def main():
    parser = argparse.ArgumentParser(description="Export pit.bin data to SQLite, Excel, or Access")
    parser.add_argument("file", nargs="?", help="Path to pit.bin file")
    parser.add_argument("-db", action="store_true", help="Export to SQLite")
    parser.add_argument("-xlsx", action="store_true", help="Export to Excel")
    parser.add_argument("-accdb", action="store_true", help="Export to Access")
    args = parser.parse_args()
    
    if not args.file:
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename(title="Select pit.bin file", filetypes=[("BIN files", "*.bin")])
        if not file_path:
            messagebox.showinfo("Error", "No file selected.")
            return
        # Create the main popup window
        root = tk.Tk()
        root.title("Export Type")
        root.geometry("300x150")

        # Add the question label
        question_label = tk.Label(root, text="Choose the export type:")
        question_label.pack(pady=10)

        db_button = tk.Button(root, text="db", command=lambda: handle_db_click(file_path, root)) 
        xlsx_button = tk.Button(root, text="xlsx", command=lambda: handle_xlsx_click(file_path, root))
        accdb_button = tk.Button(root, text="accdb", command=lambda: handle_accdb_click(file_path, root))

        # Pack the buttons into the popup window
        db_button.pack(pady=5)
        xlsx_button.pack(pady=5)
        accdb_button.pack(pady=5)

        # Run the main event loop
        root.mainloop()
        args.file = file_path
    

    entries = parse_pit_file(args.file)
    print("hello")
    if args.db:
        export_to_sqlite(entries)
    elif args.xlsx:
        export_to_excel(entries)
    elif args.accdb:
        export_to_access(entries)
    else:
        print("No export type selected. Use -db, -xlsx, or -accdb.")

if __name__ == "__main__":
    main()