import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import os
import warnings
import sys
warnings.simplefilter("ignore", UserWarning)

from openpyxl import Workbook
from openpyxl.styles import Border, Side

df = None
df_filtered = None

# =========================
# Excel automatisch laden
# =========================
def load_excel_auto():
    global df

    if getattr(sys, 'frozen', False):
        script_map = os.path.dirname(sys.executable)
    else:
        script_map = os.path.dirname(os.path.abspath(__file__))

    try:
        files = [f for f in os.listdir(script_map) if f.endswith(".xlsx")]

        if not files:
            messagebox.showerror("Fout", "Geen Excel bestand gevonden!")
            return

        file_path = os.path.join(script_map, files[0])

        df = pd.read_excel(file_path)

        if 'Location' not in df.columns or 'Location Status' not in df.columns:
            messagebox.showerror("Fout", "Je hebt je vorige bestand nog niet verwijderd!")
            return

        process_data()

        status_label.config(text=f"Loaded: {files[0]}")

    except Exception as e:
        messagebox.showerror("Fout", str(e))


# =========================
# Data verwerken
# =========================
def process_data():
    global df, df_filtered

    df['Location'] = df['Location'].astype(str)

    df['Gang'] = df['Location'].str[:3]
    df['LocatieNummer'] = df['Location'].str[3:6]

    df = df[df['LocatieNummer'].str.isnumeric()].copy()
    df['LocatieNummer'] = df['LocatieNummer'].astype(int)

    df_filtered = df[df['Location'].str[0].isin(['A', 'B', 'C'])]

    df_filtered = df_filtered[
        df_filtered['Location Status'] == "E - Empty"
    ]

    update_gangen()


# =========================
# Gangen vullen
# =========================
def update_gangen():
    gangen = sorted(df_filtered['Gang'].unique())
    gang_dropdown['values'] = gangen

    if gangen:
        gang_var.set(gangen[0])


# =========================
# TYPE DETECTIE
# =========================
def get_location_type(loc):
    loc = str(loc).strip()

    if loc.startswith("CCX"):
        return "CCX"
    elif loc.startswith("A"):
        return "A"
    elif loc.startswith("B"):
        return "B"
    elif loc.startswith("C"):
        return "C"
    else:
        return "OTHER"


# =========================
# Filter + EXPORT
# =========================
def filter_and_export():
    if df_filtered is None:
        messagebox.showwarning("Let op", "Geen data geladen")
        return

    gang = gang_var.get()
    keuze = keuze_var.get()

    # =========================
    # EXCLUDE LISTS
    # =========================
    exclude_locations_BC = [
        5, 6, 13, 14, 21, 22, 29, 30, 37, 38, 45, 46, 53, 54, 61, 62,
        69, 70, 77, 78, 85, 86, 93, 94, 101, 102, 109, 110, 117, 118,
        125, 126, 133, 134, 141, 142, 149, 150, 157, 158, 165, 166,
        173, 174, 181, 182, 189, 190, 197, 198
    ]

    exclude_locations_A = [
        9, 10, 17, 18, 25, 26, 33, 34, 41, 42, 49, 50,
        57, 58, 65, 66, 73, 74, 81, 82, 89, 90, 97, 98,
        105, 106, 113, 114, 121, 122, 129, 130, 137, 138,
        145, 146, 153, 154, 161, 162, 169, 170, 177, 178,
        185, 186, 193, 194, 201, 202, 209, 210
    ]
    
    exclude_locations_BBX = [3, 7, 11, 15, 19, 23, 27, 31, 35, 39, 43, 47, 51, 55, 59, 63, 67, 71, 75, 79]

    # =========================
    # EXCLUDE LOGICA (CORRECT)
    # =========================
    def apply_exclude(row):
        loc = str(row['Location']).strip()
        loc_type = get_location_type(loc)
        loc_num = row['LocatieNummer']

        if loc_type == "A":
            return loc_num not in exclude_locations_A

        elif loc_type == "CCX":
            return loc_num not in exclude_locations_BBX

        elif loc_type in ["B", "C"]:
            return loc_num not in exclude_locations_BC

        else:
            return False

    # =========================
    # DATA VOOR LIJST
    # =========================
    data_list = df_filtered[df_filtered['Gang'] == gang]

    if keuze == "Even":
        data_list = data_list[data_list['LocatieNummer'] % 2 == 0]
    elif keuze == "Oneven":
        data_list = data_list[data_list['LocatieNummer'] % 2 == 1]

    data_list = data_list[data_list.apply(apply_exclude, axis=1)]
    data_list = data_list.sort_values('LocatieNummer').reset_index(drop=True)
    data_list = data_list[~((data_list['LocatieNummer'] // 2) % 4 == 3)]

    if data_list.empty:
        messagebox.showwarning("Leeg", "Geen resultaten")
        return

    # =========================
    # DATA VOOR OVERZICHT
    # =========================
    data_all = df_filtered.copy()

    if keuze == "Even":
        data_all = data_all[data_all['LocatieNummer'] % 2 == 0]
    elif keuze == "Oneven":
        data_all = data_all[data_all['LocatieNummer'] % 2 == 1]

    data_all = data_all[data_all.apply(apply_exclude, axis=1)]
    data_all = data_all[~((data_all['LocatieNummer'] // 2) % 4 == 3)]

    # =========================
    # DATA VOOR FULL JUK 3
    # =========================
    data_full = df.copy()
    data_full = data_full[data_full['Location Status'] == "F - Full"]
    data_full = data_full[~((data_full['LocatieNummer'] // 2) % 4 == 3)]
    data_full = data_full[data_full['Location'].str[0].isin(['A', 'B', 'C'])]

    def apply_full_filter(row):
        loc = str(row['Location']).strip()
        loc_type = get_location_type(loc)
        loc_num = row['LocatieNummer']

        if loc_type == "A":
            return loc_num in exclude_locations_A
        elif loc_type == "CCX":
            return loc_num in exclude_locations_BBX
        elif loc_type in ["B", "C"]:
            return loc_num in exclude_locations_BC
        else:
            return False

    data_full = data_full[data_full.apply(apply_full_filter, axis=1)]
    data_full = data_full.sort_values(['Gang', 'LocatieNummer']).reset_index(drop=True)

    # =========================
    # EXPORT
    # =========================
    try:
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))

        bestandsnaam = f"Lege_locaties_{gang}_{keuze}.xlsx"
        output_path = os.path.join(base_path, bestandsnaam)

        wb = Workbook()

        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # SHEET 1
        ws1 = wb.active
        ws1.title = "Lijst"
        ws1.append(["Lege Locaties", "Afgevinkt"])

        for _, row in data_list.iterrows():
            ws1.append([row['Location'], ""])

        for row in ws1.iter_rows():
            for cell in row:
                cell.border = thin_border

        # SHEET 2
        ws2 = wb.create_sheet(title="Overzicht")

        if 'Height' in data_all.columns:
            pivot = pd.pivot_table(
                data_all,
                index='Gang',
                columns='Height',
                values='Location',
                aggfunc='count',
                fill_value=0
            ).reset_index()

            ws2.append(list(pivot.columns))

            for _, row in pivot.iterrows():
                ws2.append(list(row))

        # SHEET 3
        ws3 = wb.create_sheet(title="Full_Juk_3")

        ws3.append(["Gang", "Locatie", "Hoogte", "FULL"])



        for _, row in data_full.iterrows():
            location = str(row['Location'])

            # Laatste 2 cijfers = hoogte
            hoogte = location[-2:] if location[-2:].isdigit() else ""

            ws3.append([
                row['Gang'],
                row['LocatieNummer'],
                hoogte,
                row['Location Status']              
            ])

       

        for ws in [ws1, ws2, ws3]:
            for row in ws.iter_rows():
                for cell in row:
                    cell.border = thin_border



        ws1.page_setup.fitToWidth = 1
        ws1.page_setup.fitToHeight = False

        ws1.page_margins.top = 0
        ws1.page_margins.bottom = 0
        ws1.page_margins.left = 1
        ws1.page_margins.right = 0
        
       
        ws2.page_setup.fitToWidth = 1
        ws2.page_setup.fitToHeight = False

        ws2.page_margins.top = 0
        ws2.page_margins.bottom = 0
        ws2.page_margins.left = 0
        ws2.page_margins.right = 0
        
        ws3.page_setup.fitToWidth = 1
        ws3.page_setup.fitToHeight = False

        ws3.page_margins.top = 0
        ws3.page_margins.bottom = 0
        ws3.page_margins.left = 1
        ws3.page_margins.right = 0

        
        wb.save(output_path)

        messagebox.showinfo("Succes", f"Bestand opgeslagen:\n{bestandsnaam}")
        
    except Exception as e:
        messagebox.showerror("Export fout", str(e))


# =========================
# GUI
# =========================
root = tk.Tk()
root.title("Locatie Tool")
root.geometry("500x250")

frame = tk.Frame(root)
frame.pack(pady=20)

gang_var = tk.StringVar()
gang_dropdown = ttk.Combobox(frame, textvariable=gang_var, width=10)
gang_dropdown.grid(row=0, column=0, padx=5)

keuze_var = tk.StringVar(value="Beide")
ttk.Combobox(
    frame,
    textvariable=keuze_var,
    values=["Even", "Oneven", "Beide"],
    width=10
).grid(row=0, column=1, padx=5)

tk.Button(frame, text="Genereer Excel", command=filter_and_export).grid(row=0, column=2, padx=5)

status_label = tk.Label(root, text="Bestand laden...")
status_label.pack(pady=10)

root.after(100, load_excel_auto)
root.mainloop()