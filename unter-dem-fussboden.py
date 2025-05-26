# -*- coding: utf-8 -*-
"""
Created on Wed May 26 00:29:40 2025
version 1.0
@author: Professional
"""

import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import os
from tkinter import Tk, filedialog
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.patches import Rectangle
from datetime import datetime
from colorama import init, Fore, Style
from tqdm import tqdm

init()

def print_colored_column(index, col_name, example_value, unit_info):
    col_text = f"{index}: {col_name} ({example_value})"
    if "no-units" in str(unit_info).lower():
        print(Fore.YELLOW + col_text + Style.RESET_ALL)
    else:
        print(col_text)

def find_temp_column(df):
    metadata_rows = {}
    for idx, row in df.iterrows():
        row_str = str(row.iloc[0])
        if 'Datenpunktadresse' in row_str:
            metadata_rows['Datenpunktadresse'] = idx
        elif 'Klartext' in row_str:
            metadata_rows['Klartext'] = idx
        elif 'Einheit' in row_str:
            metadata_rows['Einheit'] = idx
        elif 'Min' in row_str:
            metadata_rows['Min'] = idx
        elif 'Max' in row_str:
            metadata_rows['Max'] = idx

    print("\nVerfügbare Datenspalten:")
    for i, col in enumerate(df.columns):
        example = str(df.iloc[metadata_rows.get('Klartext', 0), i]) if 'Klartext' in metadata_rows else str(df.iloc[0, i])
        unit_info = str(df.iloc[metadata_rows.get('Einheit', 0), i]) if 'Einheit' in metadata_rows else ""
        print_colored_column(i, col, example[:50] + ('...' if len(example) > 50 else ''), unit_info)

    while True:
        try:
            temp_col = int(input("\nBitte wählen Sie die Spaltennummer für die Temperaturvisualisierung: "))
            if 0 <= temp_col < len(df.columns):
                if all(key in metadata_rows for key in ['Datenpunktadresse', 'Klartext', 'Einheit', 'Min', 'Max']):
                    print("\nBeispieldaten:")
                    print(df.iloc[[metadata_rows['Datenpunktadresse'], 
                                 metadata_rows['Klartext'], 
                                 metadata_rows['Einheit'], 
                                 metadata_rows['Min'], 
                                 metadata_rows['Max']], [0, temp_col]])
                return temp_col
            print("Bitte geben Sie eine Nummer aus der Liste ein")
        except ValueError:
            print("Bitte geben Sie eine gültige Zahl ein")

def ask_annotation_preference():
    while True:
        response = input("\nNumerische Temperaturwerte anzeigen? (J/N): ").strip().upper()
        if response in ['J', 'N']:
            return response == 'J'
        print("Bitte J (ja) oder N (nein) eingeben")

def load_file():
    root = Tk()
    root.attributes('-topmost', True)
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Datei auswählen",
        filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv"), ("All files", "*.*")]
    )
    root.destroy()
    return file_path

def build_month_row_map(df):
    month_map = {}
    current_month = None
    start_idx = None
    
    for idx, row in df.iterrows():
        dt = row['DATETIME']
        if pd.isna(dt):
            continue
        
        month = dt.month
        if current_month is None:
            current_month = month
            start_idx = idx
        elif month != current_month:
            month_map[current_month] = (start_idx + 1, idx)
            current_month = month
            start_idx = idx

    if current_month is not None and start_idx is not None:
        month_map[current_month] = (start_idx + 1, df.index[-1] + 1)

    return month_map

def select_data_range_by_month(df):
    df = df.copy()
    df['DATETIME'] = pd.to_datetime(df.iloc[:, 0], format='%Y.%m.%d %H:%M:%S', errors='coerce')
    df = df[df['DATETIME'].notna()]
    
    month_map = build_month_row_map(df)
    if not month_map:
        print("❌ Konnte die Monatsübersicht nicht erstellen. Es werden alle Daten verwendet.")
        return 1, len(df)

    print("\n📅 Erkannte Monate im Datensatz:")
    for month, (start, end) in month_map.items():
        print(f"Monat {month:02d}: Zeilen {start} – {end}")
    
    while True:
        user_input = input(
            "\nBitte geben Sie eine Monatsnummer (1–12), eine Startzeile (innerhalb eines Monats), "
            "oder einen gültigen Bereich im Format start-end ein: ").strip()

        if user_input.isdigit():
            month_num = int(user_input)
            if month_num in month_map:
                start, end = month_map[month_num]
                print(f"✅ Monat {month_num:02d} ausgewählt, Zeilen {start} – {end}")
                return start, end
            else:
                print("❌ Ungültige Monatsnummer.")

        elif '-' in user_input:
            try:
                start_str, end_str = user_input.split('-')
                start = int(start_str)
                end = int(end_str)

                if start >= end:
                    print("❌ Die Startzeile muss kleiner als die Endzeile sein.")
                    continue

                month_start = df.loc[start-1, 'DATETIME'].month
                month_end = df.loc[end-1, 'DATETIME'].month

                if month_start != month_end:
                    print("❌ Der Bereich darf sich nur auf einen Monat beziehen.")
                    continue

                print(f"✅ Bereich ausgewählt: Zeilen {start} – {end}")
                return start, end

            except Exception:
                print("❌ Ungültiges Format. Beispiel: 3500-4500")

        else:
            try:
                start = int(user_input)
                if start < 1 or start > df.index[-1] + 1:
                    print("❌ Zeilennummer außerhalb des gültigen Bereichs.")
                    continue

                month_start = df.loc[start-1, 'DATETIME'].month
                _, end = month_map.get(month_start, (None, None))
                if start > end:
                    print("❌ Die Startzeile liegt außerhalb des Monatsbereichs.")
                    continue

                print(f"✅ Startzeile {start} gewählt (bis Ende des Monats, Zeile {end})")
                return start, end

            except:
                print("❌ Ungültige Eingabe.")


def process_data(df, temp_col, start_row, end_row):
    print("\nDaten werden verarbeitet...")
    df = df.iloc[start_row-1:end_row].copy()
    df.columns = [str(col) for col in df.columns]
    
    datetime_col = df.columns[0]
    df['DATETIME'] = pd.to_datetime(df[datetime_col], format='%Y.%m.%d %H:%M:%S', errors='coerce')
    df = df[df['DATETIME'].notna()]
    
    df['Date'] = df['DATETIME'].dt.date
    df['Weekday'] = df['DATETIME'].dt.day_name('de_DE')
    df['Hour'] = df['DATETIME'].dt.hour
    df['Minute'] = df['DATETIME'].dt.minute
    df['TimeSlot'] = (df['Hour'] * 4) + (df['Minute'] // 15)
    df['IsWeekend'] = df['DATETIME'].dt.dayofweek >= 5
    
    return df

def create_15min_plot(data, temp_col_name, show_annotations, metadata, col_num):
    print("\nDiagramm wird erstellt...")
    cmap = LinearSegmentedColormap.from_list("temp_map", ["#2b83ba", "#ffffbf", "#d7191c"])
    vmin = data[temp_col_name].min()
    vmax = data[temp_col_name].max()
    
    dates = sorted(data['Date'].unique())
    weekdays = [data[data['Date'] == date]['Weekday'].iloc[0] for date in dates]
    is_weekend = [data[data['Date'] == date]['IsWeekend'].iloc[0] for date in dates]
    time_slots = range(0, 24*4)
    
    fig, ax = plt.subplots(figsize=(24, 12))
    
    if show_annotations:
        annotation_matrix = np.empty((len(dates), 24*4), dtype=object)
    
    for date_idx, date in enumerate(tqdm(dates, desc="Tage verarbeiten")):
        date_data = data[data['Date'] == date]
        
        for slot in time_slots:
            slot_data = date_data[date_data['TimeSlot'] == slot]
            if not slot_data.empty:
                temp = slot_data[temp_col_name].mean()
                color = cmap((temp - vmin) / (vmax - vmin))
                if show_annotations:
                    annotation_matrix[date_idx, slot] = f"{temp:.1f}"
            else:
                color = 'black'
                if show_annotations:
                    annotation_matrix[date_idx, slot] = ""
            
            rect = Rectangle((date_idx, slot), 1, 1, facecolor=color, edgecolor='white', linewidth=0.3)
            ax.add_patch(rect)
    
    if show_annotations:
        print("\nAnnotationen hinzufügen...")
        for date_idx in tqdm(range(len(dates)), desc="Annotationen"):
            for slot in time_slots:
                if annotation_matrix[date_idx, slot]:
                    ax.text(
                        date_idx + 0.5, slot + 0.5,
                        annotation_matrix[date_idx, slot],
                        ha='center', va='center',
                        fontsize=6,
                        color='black' if cmap(0.5)[0] > 0.5 else 'white'
                    )

    ax.set_xlim(0, len(dates))
    ax.set_ylim(0, 24*4)
    ax.set_xticks(np.arange(len(dates)) + 0.5)
    
    labels = [f"{weekdays[i]}\n{dates[i].strftime('%d.%m')}" for i in range(len(dates))]
    xticklabels = ax.set_xticklabels(labels, rotation=45)
    
    for i, label in enumerate(xticklabels):
        if is_weekend[i]:
            label.set_color('red')
            label.set_fontweight('bold')
    
    hour_ticks = [i*4 for i in range(25)]
    ax.set_yticks(hour_ticks)
    ax.set_yticklabels([f"{h:02d}:00" for h in range(25)])
    
    for y in range(0, 24*4, 4):
        ax.axhline(y, color='gray', linestyle='-', linewidth=0.5, alpha=0.3)
    
    title_parts = []
    if 'Datenpunktadresse' in metadata:
        title_parts.append(f"Datenpunkt: {metadata['Datenpunktadresse']}")
    if 'Klartext' in metadata:
        title_parts.append(f"Beschreibung: {metadata['Klartext']}")
    if 'Einheit' in metadata:
        title_parts.append(f"Einheit: {metadata['Einheit']}")
    if 'Min' in metadata and 'Max' in metadata:
        title_parts.append(f"Bereich: {metadata['Min']} - {metadata['Max']} °C")
    
    plt.title("\n".join(title_parts), fontsize=12, pad=20)
    plt.ylabel("Uhrzeit", fontsize=12)
    
    sm = plt.cm.ScalarMappable(cmap=cmap, norm=plt.Normalize(vmin=vmin, vmax=vmax))
    sm.set_array([])
    cbar = plt.colorbar(sm, ax=ax, pad=0.02)
    cbar.set_label('Temperatur (°C)', fontsize=12)
    
    plt.tight_layout()
    return fig

def main():
    print("="*60)
    print("Fußbodenheizung Temperaturvisualisierung")
    print("="*60)
    
    file_path = load_file()
    if not file_path:
        print("Keine Datei ausgewählt. Programm wird beendet.")
        return
    
    try:
        print("\nDatei wird gelesen...")
        try:
            df = pd.read_excel(file_path)
        except:
            df = pd.read_csv(file_path)
        
        print("\nErste 5 Zeilen:")
        print(df.head())
    except Exception as e:
        print(f"\nFehler beim Lesen der Datei: {e}")
        return
    
    temp_col_idx = find_temp_column(df)
    temp_col_name = df.columns[temp_col_idx]
    
    metadata_rows = {}
    for idx, row in df.iterrows():
        row_str = str(row.iloc[0])
        if 'Datenpunktadresse' in row_str:
            metadata_rows['Datenpunktadresse'] = str(row.iloc[temp_col_idx])
        elif 'Klartext' in row_str:
            metadata_rows['Klartext'] = str(row.iloc[temp_col_idx])
        elif 'Einheit' in row_str:
            metadata_rows['Einheit'] = str(row.iloc[temp_col_idx])
        elif 'Min' in row_str:
            metadata_rows['Min'] = str(row.iloc[temp_col_idx])
        elif 'Max' in row_str:
            metadata_rows['Max'] = str(row.iloc[temp_col_idx])
    
    start_row, end_row = select_data_range_by_month(df)
    data = process_data(df, temp_col_idx, start_row, end_row)
    show_annotations = ask_annotation_preference()
    
    output_file = f"fussbodenheizung_temperatur_{temp_col_idx}.png"
    fig = create_15min_plot(data, temp_col_name, show_annotations, metadata_rows, temp_col_idx)
    
    print("\nDiagramm wird gespeichert...")
    fig.savefig(output_file, dpi=300, bbox_inches='tight')
    plt.close(fig)
    print(f"\nDiagramm gespeichert als: {output_file}")
    
    print("\nVerarbeitung erfolgreich abgeschlossen!")

if __name__ == "__main__":
    main()
