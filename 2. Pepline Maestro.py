# CODIGO_FINAL_CORREGIDO_v2.py
# -*- coding: utf-8 -*-
"""
PIPELINE MAESTRO CORREGIDO v2
- IGC: índice estático (estado actual)
- Vida dinámica: FClima x FGeot  (solo factores ambientales)
- HDM-4 tradicional: se mantiene para calibración/comparación
- Modelo propuesto: integrado (entrena con N_falla_HDM4 si no hay observaciones)
"""

import os
import re
import math
import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from scipy.stats import pearsonr
import statsmodels.api as sm
from sklearn.metrics import mean_absolute_error, mean_squared_error

import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows

# -------------------------
# CONSTANTES Y CURVAS (PCI)
# -------------------------
DV_CURVES = {
    'alligator cracking': {'L': [(0,0),(0.5,2),(1,6),(2,15),(4,35),(8,60),(15,85),(30,98)], 'M': [(0,0),(0.5,4),(1,12),(2,30),(4,55),(8,75),(15,90),(30,99)], 'H': [(0,0),(0.5,8),(1,20),(2,45),(4,70),(8,88),(15,96),(30,100)]},
    'linear cracking': {'L': [(0,0),(1,1),(3,6),(6,18),(12,40),(20,65),(30,85)], 'M': [(0,0),(1,3),(3,12),(6,30),(12,55),(20,80),(30,97)], 'H': [(0,0),(1,6),(3,20),(6,45),(12,72),(20,92),(30,100)]},
    'pothole': {'L': [(0,0),(0.1,2),(0.5,8),(1,20),(2,40),(5,70),(10,95)], 'M': [(0,0),(0.1,4),(0.5,12),(1,30),(2,60),(5,85),(10,99)], 'H': [(0,0),(0.1,8),(0.5,25),(1,50),(2,80),(5,95),(10,100)]},
    'patching': {'L': [(0,0),(0.5,2),(1,8),(2,20),(5,45),(10,78),(20,96)], 'M': [(0,0),(0.5,4),(1,15),(2,35),(5,65),(10,88),(20,99)], 'H': [(0,0),(0.5,8),(1,24),(2,50),(5,85),(10,97),(20,100)]},
    'raveling': {'L': [(0,0),(1,3),(3,9),(6,22),(12,45),(25,80)], 'M': [(0,0),(1,6),(3,18),(6,40),(12,70),(25,95)], 'H': [(0,0),(1,12),(3,30),(6,60),(12,88),(25,100)]},
    'bleeding': {'L': [(0,0),(1,2),(3,8),(6,18),(12,40),(25,70)], 'M': [(0,0),(1,4),(3,12),(6,30),(12,60),(25,88)], 'H': [(0,0),(1,8),(3,25),(6,55),(12,82),(25,100)]},
    'rutting': {'L': [(0,0),(0.5,2),(1,6),(2,18),(4,40),(8,70)], 'M': [(0,0),(0.5,4),(1,12),(2,36),(4,65),(8,90)], 'H': [(0,0),(0.5,8),(1,25),(2,55),(4,85),(8,100)]},
    'depression': {'L': [(0,0),(0.5,2),(1,8),(2,20),(4,45),(8,80)], 'M': [(0,0),(0.5,4),(1,14),(2,36),(4,70),(8,95)], 'H': [(0,0),(0.5,8),(1,28),(2,60),(4,88),(8,100)]},
    'default': {'L': [(0,0),(1,1),(3,6),(6,18),(12,40),(20,65),(30,85)], 'M': [(0,0),(1,3),(3,12),(6,30),(12,55),(20,80),(30,97)], 'H': [(0,0),(1,6),(3,20),(6,45),(12,72),(20,92),(30,100)]}
}

CDV_TABLE = {
    0:[0,0,0,0,0,0,0], 5:[5,5,5,5,5,5,5], 10:[10,10,10,10,10,10,10], 20:[20,18,18,18,18,18,18],
    30:[30,26,25,24,23,22,21], 40:[40,36,34,33,32,31,30], 50:[50,46,44,42,41,40,39], 60:[60,56,54,52,50,49,48],
    70:[70,66,64,62,60,59,58], 80:[80,76,74,72,70,69,68], 90:[90,86,84,82,80,79,78], 100:[100,96,94,92,90,89,88]
}

# Parámetros de base (se mantienen y se pueden calibrar en ejecución)
PCI_FALLA = 19.13             # umbral de falla (IGC original). Se mantiene para cálculos de N_falla.
K_ESTANDAR = 2.0e-6           # coeficiente base (se puede calibrar)
IRI_TERMINAL_HDM4 = 8.0       # terminal IRI (HDM-4)
K_HDM4_IRI = 3.0e-5           # progresión IRI (HDM-4 simplificado)

# Pesos AHP originales (se mantienen para IGC estático en informe)
W_AHP = {
    'PCI': 0.4954,
    'IRI': 0.2417,
    'IMDA': 0.1251,
    'Clima': 0.0876,
    'Suelo': 0.0502
}

# ---- funciones utilitarias ----
def interp_piecewise(points, x):
    pts = sorted(points, key=lambda p: p[0])
    xs, ys = [p[0] for p in pts], [p[1] for p in pts]
    if x <= xs[0]: return ys[0]
    if x >= xs[-1]: return ys[-1]
    for i in range(len(xs)-1):
        if xs[i] <= x <= xs[i+1]:
            x0,y0,x1,y1 = xs[i],ys[i],xs[i+1],ys[i+1]
            t=(x-x0)/(x1-x0) if (x1-x0)!=0 else 0
            return y0+t*(y1-y0)
    return ys[-1]

def get_dv_for_distress(distress_name, severity, density_pct):
    name = str(distress_name).lower()
    sev = str(severity).strip().upper()
    if sev in ['B','L','LOW','LO']: s='L'
    elif sev in ['M','MED','MEDIUM']: s='M'
    elif sev in ['A','H','HIGH']: s='H'
    else: s='M'
    mapping = {
        'alligator':'alligator cracking','fatigue':'alligator cracking', 'allig':'alligator cracking','longitudinal':'linear cracking',
        'transverse':'linear cracking','crack':'linear cracking', 'linear':'linear cracking','pothole':'pothole','rut':'rutting',
        'patch':'patching','ravel':'raveling','bleed':'bleeding', 'depress':'depression'
    }
    key=None
    for k,v in mapping.items():
        if k in name: key=v; break
    if not key: key='default'
    curve=DV_CURVES.get(key,DV_CURVES['default'])
    points=curve.get(s,curve['M'])
    return float(interp_piecewise(points,density_pct))

def compute_cdv_from_tdv_q(tdv,q):
    col = 6 if q>=6 else int(q)
    keys=sorted(CDV_TABLE.keys())
    if tdv<=keys[0]: return CDV_TABLE[keys[0]][col]
    if tdv>=keys[-1]: return CDV_TABLE[keys[-1]][col]
    for i in range(len(keys)-1):
        if keys[i]<=tdv<=keys[i+1]:
            k0,k1=keys[i],keys[i+1]
            y0,y1=CDV_TABLE[k0][col],CDV_TABLE[k1][col]
            t=(tdv-k0)/(k1-k0) if (k1-k0)!=0 else 0
            return y0+t*(y1-y0)
    return CDV_TABLE[keys[-1]][col]

def normalize_columns_pci(df):
    df=df.rename(columns=lambda c:str(c).strip())
    lower_map={c.lower():c for c in df.columns}
    mapping={}
    if 'unidad' in lower_map: mapping[lower_map['unidad']]='Unidad'
    elif 'unit' in lower_map: mapping[lower_map['unit']]='Unidad'
    if 'Unidad' in mapping.values():
        df[mapping['Unidad']]=df[mapping['Unidad']].astype(str).str.strip()
    if 'progresiva' in lower_map:
        mapping[lower_map['progresiva']]='Progresiva'
        df[mapping['Progresiva']]=df[mapping['Progresiva']].astype(str).str.strip()
    if 'área' in lower_map: mapping[lower_map['área']]='Area_m2'
    elif 'area' in lower_map: mapping[lower_map['area']]='Area_m2'
    if 'falla' in lower_map: mapping[lower_map['falla']]='Falla'
    elif 'distress' in lower_map: mapping[lower_map['distress']]='Falla'
    if 'severidad' in lower_map: mapping[lower_map['severidad']]='Severidad'
    elif 'severity' in lower_map: mapping[lower_map['severity']]='Severidad'
    if 'cantidad' in lower_map: mapping[lower_map['cantidad']]='Cantidad'
    elif 'quantity' in lower_map: mapping[lower_map['quantity']]='Cantidad'
    if 'area_falla' in lower_map: mapping[lower_map['area_falla']]='Area_Falla_m2'
    df=df.rename(columns=mapping)
    return df
# ====================================================================
# === FUNCIONES PCI, IRI y GRAFICOS ==================================
# ====================================================================

def calcular_pci_astm(path_input, unit_area_default=700.0, out_folder=None, ask_folder_if_none=True):
    p=Path(path_input)
    if not p.exists(): raise FileNotFoundError(path_input)
    if p.suffix.lower() in ['.xls','.xlsx']: df=pd.read_excel(p)
    elif p.suffix.lower()=='.csv': df=pd.read_csv(p)
    elif p.suffix.lower()=='.txt': df=pd.read_csv(p,sep='\t')
    else: raise ValueError("Unsupported file extension")
    df=normalize_columns_pci(df)
    if 'Area_m2' not in df.columns: df['Area_m2']=unit_area_default
    df['Cantidad']=pd.to_numeric(df['Cantidad'],errors='coerce').fillna(0.0)
    df['Area_m2']=pd.to_numeric(df['Area_m2'],errors='coerce').fillna(unit_area_default)
    df['Severidad']=df['Severidad'].astype(str).str.strip()
    df['Falla']=df['Falla'].astype(str)
    detalle,resumen=[],[]
    use_area_falla='Area_Falla_m2' in df.columns
    ocurrencias = {}
    for unidad, sub in df.groupby('Unidad'):
        area_unit=float(sub['Area_m2'].iloc[0]) if 'Area_m2' in sub.columns else unit_area_default
        grouped=sub.groupby(['Falla','Severidad'])
        dv_list,dv_records=[],[]
        for (falla,sev),rows in grouped:
            ocurrencias[falla]=ocurrencias.get(falla,0)+len(rows)
            area_affected=0.0
            if use_area_falla:
                area_affected=rows['Area_Falla_m2'].sum()
            else:
                for _,r in rows.iterrows():
                    q=float(r['Cantidad']); name=falla.lower()
                    if any(k in name for k in ['alligator','allig','fatigue']):
                        area_piece=q if q>=0.1 else q*0.5
                    elif any(k in name for k in ['pothole','hole']):
                        area_piece=q*0.6
                    elif any(k in name for k in ['longitudinal','transverse','crack','linear','fisura']):
                        sev0=str(sev).upper()
                        width=0.005 if sev0.startswith('B') or sev0.startswith('L') else (0.02 if sev0.startswith('M') else 0.05)
                        area_piece=q*width
                    elif 'rut' in name:
                        area_piece=q*0.5
                    else:
                        area_piece=q if q>=0.1 else q*0.5
                    area_affected+=area_piece
            density_pct=100.0*area_affected/area_unit if area_unit>0 else 0.0
            dv=get_dv_for_distress(falla,sev,density_pct)
            dv_list.append(dv)
            dv_records.append({'Unidad':unidad,'Falla':falla,'Severidad':sev,
                               'Area_affected_m2':area_affected,'Density_%':density_pct,'DV':dv})
        tdv=sum(dv_list)
        q=sum(1 for x in dv_list if x>5)
        cdv=compute_cdv_from_tdv_q(tdv,q)
        pci_unit=max(0.0,100.0-cdv)
        resumen.append({'Unidad':unidad,'TDV':round(tdv,4),'q':int(q),'CDV':round(cdv,4),'PCI':round(pci_unit,4)})
        detalle.extend(dv_records)
    df_resumen=pd.DataFrame(resumen).sort_values('Unidad')
    df_detalle=pd.DataFrame(detalle)
    pci_promedio=df_resumen['PCI'].mean() if not df_resumen.empty else np.nan
    out_folder=Path(out_folder)
    out_folder.mkdir(parents=True,exist_ok=True)
    plots_folder=out_folder/"plots_pci"
    plots_folder.mkdir(parents=True,exist_ok=True)
    timestamp=datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    excel_path=out_folder/f"Resultados_PCI_ASTMD6433_{timestamp}.xlsx"
    with pd.ExcelWriter(excel_path,engine='openpyxl') as writer:
        df_resumen.to_excel(writer,sheet_name='Resumen_PCI',index=False)
        df_detalle.to_excel(writer,sheet_name='Detalle_DV',index=False)
        stats = pd.DataFrame({'PCI_promedio':[pci_promedio],'Unidades': [len(df_resumen)]})
        stats.to_excel(writer,sheet_name='Estadisticas',index=False)
        occ_df = pd.DataFrame(sorted(ocurrencias.items(), key=lambda x: x[1], reverse=True), columns=['Falla','Ocurrencias'])
        occ_df.to_excel(writer,sheet_name='Ocurrencias',index=False)
    try:
        if ocurrencias:
            occ_items = sorted(ocurrencias.items(), key=lambda x:x[1], reverse=True)
            names=[i[0] for i in occ_items]; counts=[i[1] for i in occ_items]
            fig,ax = plt.subplots(figsize=(10,6))
            ax.barh(range(len(names)), counts)
            ax.set_yticks(range(len(names))); ax.set_yticklabels(names)
            ax.invert_yaxis(); ax.set_xlabel('Número de ocurrencias'); ax.set_title('Ocurrencias por tipo de falla')
            plt.tight_layout()
            fig.savefig(plots_folder/'ocurrencias_fallas.png', dpi=300); plt.close(fig)
    except Exception as e: print(f'Error graficando ocurrencias: {e}')
    try:
        def pci_category(p):
            if p>=85: return 'Excelente (85-100)'
            if p>=70: return 'Bueno (70-84)'
            if p>=55: return 'Regular (55-69)'
            if p>=40: return 'Malo (40-54)'
            if p>=25: return 'Muy malo (25-39)'
            return 'Fallado (0-24)'
        df_resumen['Categoria_PCI'] = df_resumen['PCI'].apply(pci_category)
        cat_counts = df_resumen['Categoria_PCI'].value_counts().sort_index()
        fig,ax = plt.subplots(figsize=(8,6))
        ax.pie(cat_counts.values, labels=cat_counts.index, autopct='%1.0f%%', startangle=90)
        ax.set_title('Distribuci\u00f3n del estado del pavimento (seg\u00fan PCI)')
        plt.tight_layout()
        fig.savefig(plots_folder/'distribucion_estado_pavimento.png', dpi=300); plt.close(fig)
    except Exception as e: print(f'Error graficando distribucion PCI: {e}')
    print(f"\n✅ PCI Calculado. Resultados en: {excel_path}")
    return {'df_resumen': df_resumen, 'excel_path': excel_path, 'plots_folder': plots_folder}

# IRI functions (sin cambios relevantes, mismas interfaces)
def read_iri_file(path):
    p = Path(path)
    if not p.exists(): raise FileNotFoundError(f"Archivo no encontrado: {path}")
    if p.suffix.lower() in ['.xls', '.xlsx']: df = pd.read_excel(p)
    elif p.suffix.lower() == '.csv': df = pd.read_csv(p)
    else: raise ValueError("Extensión de archivo IRI no compatible.")
    df.columns = df.columns.str.strip()
    return df

def aggregate_iri_by_segment(df):
    df.columns = (
        df.columns.str.strip().str.lower().str.replace(" ", "_").str.replace("-", "_")
    )
    seg_col = next((c for c in df.columns if "segment_id" in c or "unidad" in c or "tramo" in c), None)
    iri_col = next((c for c in df.columns if ("iri" in c and "m" in c) or ("iri" in c and "km" in c) or c == 'iri'), None)
    vel_col = next((c for c in df.columns if "vel" in c or "speed" in c), None)
    if seg_col is None: raise ValueError("No se encontró una columna compatible para 'Segment_ID' (debe ser 'Segment_ID', 'Unidad' o 'Tramo').")
    if iri_col is None: raise ValueError("No se encontró una columna compatible para 'IRI'.")
    df[iri_col] = pd.to_numeric(df[iri_col], errors='coerce')
    df[seg_col] = df[seg_col].astype(str).str.strip()
    agg_dict = {iri_col: 'mean'}
    if vel_col:
        df[vel_col] = pd.to_numeric(df[vel_col], errors='coerce')
        agg_dict[vel_col] = 'mean'
    agg = df.groupby(seg_col).agg(agg_dict).reset_index()
    agg = agg.rename(columns={agg.columns[0]: "Unidad"})
    new_cols_iri = ["IRI_m_per_km"]
    new_cols_vel = ([ "Velocidad_mean_kmh"] if vel_col and len(agg.columns)>2 else [])
    rename_dict = {}
    if iri_col in agg.columns:
        rename_dict[iri_col] = new_cols_iri[0]
    if vel_col and len(new_cols_vel) > 0 and vel_col in agg.columns:
        rename_dict[vel_col] = new_cols_vel[0]
    agg = agg.rename(columns=rename_dict)
    agg['Unidad'] = agg['Unidad'].astype(str)
    agg = agg.dropna(subset=['Unidad'])
    return agg

def process_iri_data_and_plot(base_out_folder):
    root = tk.Tk(); root.withdraw()
    num_recorridos = simpledialog.askinteger(
        "Número de Recorridos IRI",
        "Ingrese el número de recorridos IRI que tiene:",
        initialvalue=3, minvalue=1, parent=root
    )
    if num_recorridos is None: return None
    out_dir = Path(base_out_folder) / "IRI_Results"
    out_dir.mkdir(parents=True, exist_ok=True)
    plots_folder = out_dir / "plots_iri"
    plots_folder.mkdir(exist_ok=True)
    iri_agg_list = []
    for i in range(1, num_recorridos + 1):
        messagebox.showinfo("Ingreso IRI", f"Selecciona el archivo IRI del recorrido #{i}")
        iri_path = filedialog.askopenfilename(
            title=f"Seleccionar IRI recorrido {i} de {num_recorridos}",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")]
        )
        if not iri_path:
            messagebox.showwarning("Advertencia", f"No se seleccionó archivo para el recorrido {i}. Cancelando.")
            return None
        try:
            agg = aggregate_iri_by_segment(read_iri_file(iri_path))
            agg = agg.rename(columns={"IRI_m_per_km": f"IRI_m_per_km_r{i}"})
            if "Velocidad_mean_kmh" in agg.columns:
                 agg = agg.rename(columns={"Velocidad_mean_kmh": f"Velocidad_r{i}"})
            agg["source_file"] = Path(iri_path).name
            iri_agg_list.append(agg)
            print(f"Recorrido #{i} agregado. Segmentos: {len(agg)}")
        except Exception as e:
            messagebox.showerror("Error de Procesamiento", f"Fallo al procesar el archivo del recorrido {i}: {e}")
            return None
    if not iri_agg_list:
        messagebox.showwarning('Advertencia', 'No se cargó ningún archivo IRI válido.')
        return None
    merged = iri_agg_list[0].copy()
    for i in range(1, len(iri_agg_list)):
        iri_agg_list[i]['Unidad'] = iri_agg_list[i]['Unidad'].astype(str)
        merged = merged.merge(iri_agg_list[i], on="Unidad", how="outer")
    cols_iri = [f"IRI_m_per_km_r{i+1}" for i in range(len(iri_agg_list))]
    cols_vel = [f"Velocidad_r{i+1}" for i in range(len(iri_agg_list)) if f"Velocidad_r{i+1}" in merged.columns]
    merged["IRI_mean_all"] = merged[cols_iri].mean(axis=1, skipna=True)
    merged["IRI_std_all"] = merged[cols_iri].std(axis=1, skipna=True)
    merged["n_runs"] = merged[cols_iri].notna().sum(axis=1)
    iri_recorrido_promedios = {f"Recorrido {i+1}": merged[col].mean(skipna=True) for i, col in enumerate(cols_iri)}
    iri_promedio_total = merged["IRI_mean_all"].mean()
    print("\n--- Resultados de Recorridos IRI ---")
    for rec, iri_mean in iri_recorrido_promedios.items():
        print(f"{rec} Promedio: {iri_mean:.3f} m/km")
    print(f"IRI Promedio Total (General): {iri_promedio_total:.3f} m/km")
    output_path = out_dir / f"Resultados_IRI_{len(iri_agg_list)}recorridos_merged.xlsx"
    merged.to_excel(output_path, index=False)
    print(f"Resultados IRI guardados en: {output_path}")
    merged['Unidad_str'] = merged['Unidad'].astype(str)
    try:
        fig, ax = plt.subplots(figsize=(12, 6))
        for col in cols_iri: ax.plot(merged['Unidad_str'], merged[col], marker='o', linestyle='-', label=col.replace('IRI_m_per_km_', ''))
        ax.plot(merged['Unidad_str'], merged['IRI_mean_all'], linestyle='--', color='black', label='IRI Promedio Total'); ax.set_title('IRI por Segmento: Comparativo de Recorridos y Promedio')
        ax.set_xlabel('Segmento (Unidad)'); ax.set_ylabel('IRI (m/km)'); ax.legend(); plt.xticks(rotation=45, ha='right'); plt.grid(axis='y', linestyle='--'); plt.tight_layout()
        fig.savefig(plots_folder / 'G1_IRI_comparativo_recorridos_lineas.png', dpi=300); plt.close(fig)
    except Exception as e: print(f"Error al generar G1 (IRI comparativo de líneas): {e}")
    try:
        rec_names = list(iri_recorrido_promedios.keys()); rec_values = list(iri_recorrido_promedios.values())
        fig, ax = plt.subplots(figsize=(10, 6))
        bars = ax.bar(rec_names, rec_values)
        ax.axhline(iri_promedio_total, color='red', linestyle='--', label=f'IRI Promedio General ({iri_promedio_total:.2f})'); ax.set_title('IRI Promedio de Cada Recorrido')
        ax.set_ylabel('IRI Promedio (m/km)'); ax.set_xlabel('Recorrido'); ax.legend()
        for bar in bars: yval = bar.get_height(); ax.text(bar.get_x() + bar.get_width()/2.0, yval, f'{yval:.3f}', va='bottom', ha='center')
        plt.tight_layout()
        fig.savefig(plots_folder / 'G2_IRI_promedio_comparativo_barras.png', dpi=300); plt.close(fig)
    except Exception as e: print(f"Error al generar G2 (IRI promedio comparativo barras): {e}")
    if cols_vel:
        try:
            fig, ax = plt.subplots(figsize=(12, 6))
            for col in cols_vel: ax.plot(merged['Unidad_str'], merged[col], marker='.', linestyle='-', label=col.replace('Velocidad_', ''))
            ax.set_title('Velocidad Promedio por Segmento (Recorridos Comparados)'); ax.set_xlabel('Segmento (Unidad)'); ax.set_ylabel('Velocidad (km/h)'); ax.legend()
            plt.xticks(rotation=45, ha='right'); plt.grid(axis='y', linestyle='--'); plt.tight_layout()
            fig.savefig(plots_folder / 'G3_Velocidad_comparativo_lineas.png', dpi=300); plt.close(fig)
        except Exception as e: print(f"Error al generar G3 (Velocidad comparativo): {e}")
    try:
        def get_rcdec_category(iri):
            if iri >= 0 and iri <= 2: return 'Excelente (0-2)'
            if iri > 2 and iri <= 4: return 'Bueno (2-4)'
            if iri > 4 and iri <= 6: return 'Regular (4-6)'
            if iri > 6 and iri <= 8: return 'Deficiente (6-8)'
            if iri > 8: return 'Crítico (>8)'
            return 'N/A'
        df_dist = merged[['Unidad', 'IRI_mean_all']].copy()
        df_dist['Categoria_Rcdec'] = df_dist['IRI_mean_all'].apply(get_rcdec_category)
        category_order = ['Excelente (0-2)', 'Bueno (2-4)', 'Regular (4-6)', 'Deficiente (6-8)', 'Crítico (>8)']
        cat_counts = df_dist['Categoria_Rcdec'].value_counts().reindex(category_order, fill_value=0)
        fig, ax = plt.subplots(figsize=(8, 8))
        ax.pie(cat_counts.values, labels=cat_counts.index, autopct='%1.1f%%', startangle=90, wedgeprops=dict(width=0.4), pctdistance=0.8)
        ax.set_title('Distribuci\u00f3n del Estado Superficial (IRI / Categor\u00eda Rcdec)')
        plt.tight_layout()
        fig.savefig(plots_folder / 'G4_Distribucion_estado_Rcdec.png', dpi=300); plt.close(fig)
    except Exception as e: print(f"Error al generar G4 (Distribuci\u00f3n Rcdec): {e}")
    df_return = merged[['Unidad', 'IRI_mean_all']].copy()
    df_return['Unidad'] = df_return['Unidad'].astype(str)
    return df_return
# ====================================================================
# === INTEGRACIÓN PCI + IRI, CORRELACIÓN y ESALs =======================
# ====================================================================

def merge_pci_iri(df_pci_resumen, df_iri_summary, base_out_folder):
    df_iri_summary = df_iri_summary.rename(columns={'IRI_mean_all': 'IRI'})
    df_pci_resumen['Unidad'] = df_pci_resumen['Unidad'].astype(str)
    df_iri_summary['Unidad'] = df_iri_summary['Unidad'].astype(str)
    df_merged = df_pci_resumen[['Unidad', 'PCI']].merge(df_iri_summary, on='Unidad', how='outer')
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    excel_path = Path(base_out_folder) / f"Resultados_Integrados_PCI_IRI_{timestamp}.xlsx"
    df_merged.to_excel(excel_path, index=False, engine='openpyxl')
    print(f"\n✅ INTEGRACIÓN PCI + IRI exitosa. Archivo guardado en: {excel_path}")
    return df_merged

def perform_pci_iri_correlation(df_merged_data, base_out_folder):
    print("\n--- PASO 4: Iniciando Análisis de Correlación PCI vs IRI ---")
    plots_folder = Path(base_out_folder) / "Correlation_Results"
    plots_folder.mkdir(exist_ok=True)
    df_corr = df_merged_data.dropna(subset=['PCI', 'IRI']).copy()
    df_corr['PCI'] = pd.to_numeric(df_corr['PCI'], errors='coerce')
    df_corr['IRI'] = pd.to_numeric(df_corr['IRI'], errors='coerce')
    df_corr = df_corr.dropna(subset=['PCI', 'IRI'])
    if len(df_corr) < 2:
        print("Advertencia: Menos de 2 muestras válidas para la correlación. Saliendo del análisis.")
        return None, None
    r, p_value = pearsonr(df_corr['IRI'], df_corr['PCI'])
    X = df_corr['IRI']
    Y = df_corr['PCI']
    X = sm.add_constant(X)
    model = sm.OLS(Y, X).fit()
    intercept = model.params['const']
    slope = model.params['IRI']
    r_squared = model.rsquared
    print(f"\nResultados Estadísticos (n = {len(df_corr)} muestras):")
    print(f"  - Coeficiente de Correlación de Pearson (r): {r:.4f}")
    print(f"  - Valor p (Significancia): {p_value:.4f}")
    print(f"  - Coeficiente de Determinación (R²): {r_squared:.4f}")
    print(f"  - Ecuación de Regresión: PCI = {slope:.3f} * IRI + {intercept:.3f}")
    try:
        fig, ax = plt.subplots(figsize=(10, 7))
        ax.scatter(df_corr['IRI'], df_corr['PCI'], alpha=0.6, label='Datos Muestreados')
        ax.plot(df_corr['IRI'], model.predict(X), color='red', linewidth=2, label=f'Regresión Lineal\nPCI = {slope:.3f} * IRI + {intercept:.3f}\nR^2 = {r_squared:.4f}')
        ax.set_title('Correlación entre PCI (Condición) e IRI (Rugosidad)')
        ax.set_xlabel('IRI Promedio (m/km)'); ax.set_ylabel('PCI'); ax.grid(True, linestyle='--', alpha=0.6); ax.legend(loc='lower left')
        plot_path = plots_folder / 'Correlacion_PCI_vs_IRI_Regresion.png'
        fig.savefig(plot_path, dpi=300); plt.close(fig)
        print(f"Gráfico de dispersión guardado en: {plot_path}")
    except Exception as e:
        print(f"Error al generar el gráfico de correlación: {e}")
    stats_summary = pd.DataFrame({
        'Métrica': ['Pearson_r', 'P_value', 'R_squared', 'Intercept', 'Slope', 'N_samples'],
        'Valor': [r, p_value, r_squared, intercept, slope, len(df_corr)]
    })
    stats_path = plots_folder / 'Resumen_Estadistico_PCI_IRI.xlsx'
    stats_summary.to_excel(stats_path, index=False, engine='openpyxl')
    print(f"Resultados estadísticos guardados en: {stats_path}")
    return model, df_corr

# Factor camión (FC) - sin cambios
def load_and_calculate_fc():
    root = tk.Tk(); root.withdraw()
    messagebox.showinfo('Tráfico - Factor Camión', 'Selecciona la tabla de clasificación vehicular (porcentajes y FEA) o cancela para usar valor por defecto.')
    fc_path = filedialog.askopenfilename(title='Seleccionar Archivo de Factor Camión (FC/FEA)', filetypes=[('Excel files','*.xls *.xlsx')])
    if not fc_path:
        messagebox.showwarning('Advertencia', 'Archivo FC no seleccionado. Usando FC por defecto (0.5627).')
        return 0.5627
    try:
        df_fc = pd.read_excel(fc_path)
        df_fc.columns = df_fc.columns.str.strip().str.lower().str.replace(' ', '_')
        perc_col = next((c for c in df_fc.columns if 'porcentaje' in c or 'pct' in c), None)
        fea_col = next((c for c in df_fc.columns if 'fea' in c), None)
        if not perc_col or not fea_col:
            raise ValueError("Las columnas 'Porcentaje' y 'FEA' no fueron encontradas.")
        df_fc[perc_col] = pd.to_numeric(df_fc[perc_col], errors='coerce').fillna(0)
        df_fc[fea_col] = pd.to_numeric(df_fc[fea_col], errors='coerce').fillna(0)
        df_fc['EAL'] = df_fc[perc_col] * df_fc[fea_col]
        fc_value = df_fc['EAL'].sum()
        print(f"  - Factor Camión (FC) calculado desde el archivo: {fc_value:.4f}")
        return fc_value
    except Exception as e:
        messagebox.showerror('Error FC', f"Fallo al calcular el Factor Camión: {e}. Usando FC por defecto (0.5627).")
        return 0.5627

def calculate_esals_interpolado(df_pci_resumen, base_out_folder):
    print("\n--- PASO 5: Iniciando Cálculo de ESALs por Metodología de Interpolación ---")
    root = tk.Tk(); root.withdraw()
    try:
        imda_max_base = simpledialog.askfloat("FASE I: Atenuación (IMDA)", "1. IMDA MÁXIMO (Punto A) en el Año del Estudio (Ej: 450):", initialvalue=450.0)
        if imda_max_base is None: return None, None, None
        atenuacion_pct = simpledialog.askfloat("FASE I: Atenuación (IMDA)", "2. Tasa de Atenuación (%) del tráfico a lo largo del tramo (Ej: 20):", initialvalue=20.0)
        if atenuacion_pct is None: return None, None, None
        imda_min_base = imda_max_base * (1 - atenuacion_pct / 100.0)
        ano_estudio = simpledialog.askinteger("FASE II: Proyección Temporal", "3. Año en que se realizó el Estudio de Tráfico (Ej: 2014):", initialvalue=2014)
        if ano_estudio is None: return None, None, None
        ano_base = simpledialog.askinteger("FASE II: Proyección Temporal", "4. Año Base de Modelado (Ej: 2025):", initialvalue=2025)
        if ano_base is None: return None, None, None
        r_proy_pct = simpledialog.askfloat("FASE II: Proyección Temporal", "5. Tasa de Crecimiento Anual (%) para la Proyección (r) (Ej: 2.5):", initialvalue=2.5)
        if r_proy_pct is None: return None, None, None
        r_proy = r_proy_pct / 100.0
        fe_factor = simpledialog.askfloat("FASE II: Proyección Temporal", "6. Factor Estacional (Fe) (Ej: 1.05):", initialvalue=1.05)
        if fe_factor is None: return None, None, None
        anos_proyeccion = ano_base - ano_estudio
        if anos_proyeccion < 0:
            messagebox.showerror("Error", "El Año Base debe ser posterior al Año del Estudio.")
            return None, None, None
        vida_util_diseno = simpledialog.askinteger("FASE IV: Diseño Final", "7. Ingrese la Vida Útil de Diseño (L) en años:", initialvalue=20)
        if vida_util_diseno is None: return None, None, None
        carriles_diseno = simpledialog.askinteger("FASE IV: Diseño Final", "8. Número de Carriles en la Dirección de Diseño:", initialvalue=1)
        if carriles_diseno is None: return None, None, None
    except Exception as e:
        messagebox.showerror("Error de Ingreso", f"Error en el ingreso de datos de tráfico: {e}")
        return None, None, None
    df_pci_seg = df_pci_resumen.copy()
    df_pci_seg['Unidad'] = df_pci_seg['Unidad'].astype(str)
    if 'Unidad' in df_pci_seg.columns:
        df_pci_seg['K'] = df_pci_seg['Unidad'].str.replace('+', '.', regex=False).str.extract(r'(\d+\.?\d*)').astype(float)
    else:
        messagebox.showerror("Error", "No se encontró columna 'Unidad' para la interpolación del IMDA.")
        return None, None, None
    df_pci_seg = df_pci_seg.dropna(subset=['K']).sort_values('K')
    if df_pci_seg.empty:
        messagebox.showerror("Error", "La columna de progresiva está vacía o es inválida después de la limpieza.")
        return None, None, None
    K_min = df_pci_seg['K'].min()
    K_max = df_pci_seg['K'].max()
    longitud = K_max - K_min
    if longitud > 0:
        df_pci_seg['IMDA_Interpolado_2014'] = (imda_min_base + (imda_max_base - imda_min_base) * ((K_max - df_pci_seg['K']) / longitud))
    else:
        df_pci_seg['IMDA_Interpolado_2014'] = imda_max_base
    df_pci_seg['IMDABase_Proyectado'] = df_pci_seg['IMDA_Interpolado_2014'] * (1 + r_proy)**anos_proyeccion * fe_factor
    IMDABase_i_promedio = df_pci_seg['IMDABase_Proyectado'].mean()
    FC = load_and_calculate_fc()
    Fd = 1.0 / carriles_diseno
    if r_proy > 0:
        FCD = ((1 + r_proy)**vida_util_diseno - 1) / r_proy
    else:
        FCD = vida_util_diseno
    df_pci_seg['ESALs_Totales_Diseno'] = (df_pci_seg['IMDABase_Proyectado'] * 365 * FCD * Fd * FC)
    df_pci_seg['ESALs_Anuales'] = (df_pci_seg['IMDABase_Proyectado'] * 365 * Fd * FC)
    Esals_Totales_Diseno_promedio = df_pci_seg['ESALs_Totales_Diseno'].mean()
    print(f"  - Factor de Crecimiento de Diseño (FCD): {FCD:.4f}")
    print(f"  - Factor de Distribución por Carril (Fd): {Fd:.4f}")
    print(f"  - ESALs Totales de Diseño Promedio: {Esals_Totales_Diseno_promedio:,.2f}")
    print("✅ ESALs calculados por segmento (IMDA interpolado).")
    esals_params = {
        'IMDA_Max_Base': imda_max_base, 'Atenuacion_Pct': atenuacion_pct, 'IMDA_Min_Base': imda_min_base,
        'Ano_Estudio': ano_estudio, 'Ano_Base': ano_base, 'Tasa_Crecimiento_r_pct': r_proy_pct,
        'Factor_Estacional_Fe': fe_factor, 'Vida_Util_L': vida_util_diseno, 'Carriles_Diseno': carriles_diseno,
        'Factor_Camion_FC': FC, 'Factor_Crecimiento_FCD': FCD, 'Factor_Distribucion_Fd': Fd,
        'IMDA_Base_Proyectado_Promedio': IMDABase_i_promedio, 'ESALs_Totales_Diseno_Promedio': Esals_Totales_Diseno_promedio
    }
    return df_pci_seg, Esals_Totales_Diseno_promedio, esals_params
# ====================================================================
# === FACTORES CLIMATICOS Y GEOTECNICOS, IGC ESTATICO, VIDA DINAMICA ==
# ====================================================================

def calculate_fclima(base_out_folder):
    """
    Calcula FClima. Los defaults han sido ajustados para zona altoandina.
    El usuario puede cambiar los coeficientes en los diálogos si lo desea.
    """
    print("\n--- PASO 6: Iniciando Cálculo del Factor Climático (FClima) ---")
    root = tk.Tk(); root.withdraw()
    messagebox.showinfo('Coeficientes Climáticos', 'Ingresar coeficientes (valores por defecto calibrados para zona altoandina).')
    try:
        # Valores por defecto calibrados (altoandino)
        c_h = simpledialog.askfloat("coef_ch (humedad)", "Coeficiente CH - Humedad (por defecto altoandino 0.00025):", initialvalue=0.00025)
        if c_h is None: c_h = 0.00025
        c_p = simpledialog.askfloat("coef_cp (base humedad)", "Coeficiente CP - Base Humedad (por defecto 1.00):", initialvalue=1.00)
        if c_p is None: c_p = 1.00
        c_t = simpledialog.askfloat("coef_ct (temperatura)", "Coeficiente CT - Temperatura (por defecto 0.025):", initialvalue=0.025)
        if c_t is None: c_t = 0.025
        c_t0 = simpledialog.askfloat("coef_ct0 (base temp)", "Coeficiente CT0 - Base Temperatura (por defecto 0.75):", initialvalue=0.75)
        if c_t0 is None: c_t0 = 0.75
    except Exception as e:
        messagebox.showerror('Error de Ingreso', f"Fallo al ingresar los coeficientes: {e}. Asumiendo FClima = 1.0.")
        return 1.0
    messagebox.showinfo('Climatología', 'Selecciona el archivo Excel con los datos históricos MENSUALES de T. Máx, T. Mín y Precipitación.')
    clima_file_path = filedialog.askopenfilename(title='Seleccionar Archivo Histórico de Clima (Mensual)', filetypes=[('Excel files','*.xls *.xlsx')])
    if not clima_file_path:
        messagebox.showwarning('Advertencia', 'Archivo de Climatología no seleccionado. Asumiendo FClima = 1.0 (Neutral).')
        return 1.0
    try:
        df_clima_raw = pd.read_excel(clima_file_path)
        df_clima_raw.columns = df_clima_raw.columns.str.strip().str.replace(' ', '_').str.replace('.', '', regex=False).str.replace('(', '', regex=False).str.replace(')', '', regex=False).str.replace('/', '_')
        mapping = {}
        for col in df_clima_raw.columns:
            if 'TEMPMAX' in col.upper(): mapping[col] = 'TempMax'
            elif 'TEMPMIN' in col.upper(): mapping[col] = 'TempMin'
            elif 'PRECIP' in col.upper() or 'PCP' in col.upper() or 'LLUVIA' in col.upper(): mapping[col] = 'Precip'
            elif 'AÑO' in col.upper() or 'YEAR' in col.upper(): mapping[col] = 'Año'
            elif 'MES' in col.upper() or 'MONTH' in col.upper(): mapping[col] = 'Mes'
        df_clima = df_clima_raw.rename(columns=mapping)
        required_data_cols = ['TempMax', 'TempMin', 'Precip']
        df_clima = df_clima.dropna(subset=required_data_cols)
        for col in required_data_cols:
            df_clima[col] = pd.to_numeric(df_clima[col], errors='coerce')
        df_clima['Temp_Media_Mensual'] = (df_clima['TempMax'] + df_clima['TempMin']) / 2.0
        T_media_avg = df_clima['Temp_Media_Mensual'].mean()
        df_anual = df_clima.groupby('Año').agg({'Precip': 'sum'}).reset_index()
        P_anual_avg = df_anual['Precip'].mean() if 'Precip' in df_anual.columns else 0.0
        F_H = max(0.01, c_p - c_h * P_anual_avg)
        F_T = max(0.01, c_t0 + c_t * T_media_avg)
        F_Clima = F_H * F_T
        print(f"  - Coeficientes usados: ch={c_h:.5f}, cp={c_p:.3f}, ct={c_t:.5f}, ct0={c_t0:.3f}")
        print(f"  - T. Media Histórica: {T_media_avg:.4f} °C | Precip. Anual Promedio: {P_anual_avg:.4f} mm/año")
        print(f"  - Factor de Humedad (F_H): {F_H:.4f} | Factor de Temperatura (F_T): {F_T:.4f}")
        print(f"  - Factor Climático (FClima): {F_Clima:.4f}")
        return F_Clima
    except Exception as e:
        messagebox.showerror('Error en FClima', f"Fallo al calcular FClima: {e}. Asumiendo FClima = 1.0 (Neutral).")
        print(f"Error en FClima: {e}")
        return 1.0

# Tabla FGeot (sin cambios en valores, se pueden ampliar)
FGEOT_TABLE = {
    'A-1-A': 1.00, 'A-1-B': 0.90, 'A-2': 0.80, 'A-3': 0.70, 'A-4': 0.50, 'A-5': 0.40, 'A-6': 0.25, 'A-7-5': 0.15, 'A-7-6': 0.10
}

def get_fgeot_from_aashto():
    """
    Pide la clasificación AASHTO de la calicata y retorna FGeot.
    Renombré los prompts para mayor claridad.
    """
    print("\n--- PASO 7: Iniciando Cálculo del Factor Geotécnico (FGeot) ---")
    root = tk.Tk(); root.withdraw()
    opciones = list(FGEOT_TABLE.keys())
    messagebox.showinfo('Factor Geotécnico', 'Ingrese la clasificación AASHTO determinada en laboratorio (ejemplo A-1-B, A-4).')
    clase_suelo = simpledialog.askstring(
        "Ingresar Clasificación AASHTO",
        f"Ingrese la clasificación AASHTO (Ej: A-1-B, A-4, etc.). Opciones comunes: {', '.join(opciones)}:",
        initialvalue='A-1-B'
    )
    if not clase_suelo:
        messagebox.showwarning('Advertencia', 'Clasificación AASHTO no ingresada. Asumiendo FGeot = 1.0.')
        return 1.0
    clase_suelo = clase_suelo.strip().upper().replace('A7-5','A-7-5').replace('A7-6','A-7-6')
    fgeot_value = FGEOT_TABLE.get(clase_suelo)
    if fgeot_value is not None:
        print(f"  - Clasificación AASHTO ingresada: {clase_suelo}")
        print(f"  - Factor Geotécnico (FGeot) asignado: {fgeot_value:.2f}")
        return fgeot_value
    else:
        messagebox.showerror('Error de Clasificación', f"Clasificación '{clase_suelo}' no encontrada en la Tabla. Asumiendo FGeot = 1.0.")
        return 1.0

# HDM-4 tradicional (sin cambios)
def project_life_hdm4_iri(df_model_input, pci_iri_model):
    df = df_model_input.copy()
    if pci_iri_model and 'IRI' in pci_iri_model.params:
        slope = pci_iri_model.params['IRI']
        intercept = pci_iri_model.params['const']
    else:
        slope, intercept = -10.00, 100.0
    pci_falla_equivalente = max(0.0, slope * IRI_TERMINAL_HDM4 + intercept)
    df['Delta_IRI_Falla'] = IRI_TERMINAL_HDM4 - df['IRI']
    df.loc[df['Delta_IRI_Falla'] <= 0, 'N_Falla_HDM4'] = 0.0
    df.loc[df['Delta_IRI_Falla'] > 0, 'N_Falla_HDM4'] = (df['Delta_IRI_Falla'] / K_HDM4_IRI)
    df['Vida_Util_HDM4_Anios'] = df['N_Falla_HDM4'] / df['ESALs_Anuales'].replace(0, np.nan)
    df['Vida_Util_HDM4_Anios'] = df['Vida_Util_HDM4_Anios'].clip(lower=0.0).fillna(0.0)
    print(f"\n--- Resultados de Proyección HDM-4 (Tradicional) ---")
    print(f"  - IRI Terminal asumido: {IRI_TERMINAL_HDM4} m/km")
    print(f"  - K_HDM4 (Progresión IRI) asumido: {K_HDM4_IRI:.2e} 1/ESALs")
    print(f"  - PCI Equivalente de Falla (Estimado): {pci_falla_equivalente:.2f}")
    print(f"  - Vida Útil HDM-4 Restante Promedio: {df['Vida_Util_HDM4_Anios'].mean():.2f} años")
    return df[['Unidad','N_Falla_HDM4','Vida_Util_HDM4_Anios']], pci_falla_equivalente

# -------------------------
# IGC estático + Vida dinámica (FClima–FGeot)
# -------------------------
def calculate_igc_and_project_pci(df_merged_pci_esals, FClima, FGeot, base_out_folder):
    """
    - Calcula IGC (estático) para diagnóstico y reporte.
    - Calcula Vida Dinámica basada SOLAMENTE en FClima x FGeot (ya no usa PCI/IRI/IMDA en el IGC).
    - Mantengo PCI_FALLA y K_ESTANDAR como parámetros (el usuario puede calibrarlos en ejecución).
    """
    print("\n--- PASO 8: IGC (estático) y Proyección de Vida Dinámica (FClima x FGeot) ---")
    df = df_merged_pci_esals.copy()
    # Normalizaciones para IGC (solo para informe)
    IRI_MAX = 8.0
    IRI_MIN = 2.0
    IMDA_MAX = df['IMDABase_Proyectado'].max() if 'IMDABase_Proyectado' in df.columns else 0.0
    IMDA_MIN = df['IMDABase_Proyectado'].min() if 'IMDABase_Proyectado' in df.columns else 0.0
    df['N_PCI'] = df['PCI'] / 100.0
    df['N_IRI'] = (IRI_MAX - df['IRI']) / (IRI_MAX - IRI_MIN)
    df['N_IRI'] = df['N_IRI'].clip(0.0, 1.0)
    if (IMDA_MAX - IMDA_MIN) > 0:
        df['N_IMDA'] = (IMDA_MAX - df['IMDABase_Proyectado']) / (IMDA_MAX - IMDA_MIN)
    else:
        df['N_IMDA'] = 1.0
    df['N_IMDA'] = df['N_IMDA'].clip(0.0, 1.0)
    df['N_Clima'] = FClima
    df['N_Suelo'] = FGeot
    # IGC para reporte: mantiene la combinación con pesos AHP (solo lectura)
    df['IGC'] = (
        W_AHP['PCI'] * df['N_PCI'] +
        W_AHP['IRI'] * df['N_IRI'] +
        W_AHP['IMDA'] * df['N_IMDA'] +
        W_AHP['Clima'] * df['N_Clima'] +
        W_AHP['Suelo'] * df['N_Suelo']
    )
    df['Clasificacion_IGC'] = df['IGC'].apply(lambda igc: 'Excelente' if igc>=0.85 else ('Satisfactorio' if igc>=0.70 else ('Regular' if igc>=0.55 else ('Malo' if igc>=0.40 else ('Serio' if igc>=0.25 else 'Grave')))))
    print(f"  - IGC Promedio del tramo (estático, informe): {df['IGC'].mean():.4f}")
    # --- VIDA DINÁMICA: usar SOLO FClima y FGeot para ajustar K ---
    # F_dinamico (multiplicativo)
    df['F_Dinamico'] = (FClima * FGeot)
    # Evitar valores extremadamente bajos/altos
    df['F_Dinamico'] = df['F_Dinamico'].clip(lower=0.01)
    # K dinámico (ajuste de deterioro por efecto ambiental)
    df['K_Dinamico'] = K_ESTANDAR * (1.0 / df['F_Dinamico'])
    # N_falla dinámico (si PCI>PCI_FALLA)
    df['PCI_Relacion'] = df['PCI'] / PCI_FALLA
    df.loc[df['PCI_Relacion'] <= 1.0, 'N_Falla_Dinamico'] = 0.0
    mask = df['PCI_Relacion'] > 1.0
    df.loc[mask, 'N_Falla_Dinamico'] = (1.0 / df.loc[mask, 'K_Dinamico']) * np.log(df.loc[mask, 'PCI_Relacion'])
    # Vida dinámica (años) = N_falla_dinamico / ESALs_Anuales
    df['Vida_Dinamica_FClima_FGeot_Anios'] = df['N_Falla_Dinamico'] / df['ESALs_Anuales'].replace(0, np.nan)
    df['Vida_Dinamica_FClima_FGeot_Anios'] = df['Vida_Dinamica_FClima_FGeot_Anios'].clip(lower=0.0).fillna(0.0)
    print(f"  - Vida Dinámica (FClima x FGeot) Promedio: {df['Vida_Dinamica_FClima_FGeot_Anios'].mean():.2f} años")
    # Gráfica ejemplo (primer segmento cercano al promedio)
    plots_folder = Path(base_out_folder) / "IGC_Proyeccion_Results"
    plots_folder.mkdir(exist_ok=True)
    if not df.empty:
        pci_promedio = df['PCI'].mean()
        idx_ejemplo = (df['PCI'] - pci_promedio).abs().argsort().iloc[0]
        segmento_ejemplo = df.loc[idx_ejemplo]
        pci_0_seg = segmento_ejemplo['PCI']
        k_dyn_seg = segmento_ejemplo['K_Dinamico']
        esals_anuales_seg = segmento_ejemplo['ESALs_Anuales']
        if esals_anuales_seg > 0 and k_dyn_seg > 0:
            max_esals = segmento_ejemplo['N_Falla_Dinamico'] * 1.5 if segmento_ejemplo['N_Falla_Dinamico']>0 else 10000
            esals_range = np.linspace(0, max_esals, 100)
            pci_proyectado = pci_0_seg * np.exp(-k_dyn_seg * esals_range)
            anos_proyectados = esals_range / esals_anuales_seg
            try:
                fig, ax = plt.subplots(figsize=(10, 6))
                ax.plot(anos_proyectados, pci_proyectado, label=f'PCI Proyectado (K_Dinamico={k_dyn_seg:.2e})')
                ax.axhline(PCI_FALLA, color='red', linestyle='--', label=f'Umbral de Falla (PCI={PCI_FALLA})')
                ax.scatter(segmento_ejemplo['Vida_Dinamica_FClima_FGeot_Anios'], PCI_FALLA, color='red', marker='o', s=100, zorder=5, label=f"Vida Dinámica Restante: {segmento_ejemplo['Vida_Dinamica_FClima_FGeot_Anios']:.2f} años")
                ax.set_title(f"G5: Proyección de Vida Dinámica (Segmento: {segmento_ejemplo['Unidad']}) - FClima x FGeot")
                ax.set_xlabel('Años a partir del Estudio Base'); ax.set_ylabel('PCI'); ax.set_ylim(0,100); ax.grid(True, linestyle='--', alpha=0.6); ax.legend(loc='upper right')
                plot_path = plots_folder / f"G5_Proyeccion_Vida_Dinamica_Segmento_{segmento_ejemplo['Unidad'].replace('+','-')}.png"
                fig.savefig(plot_path, dpi=300); plt.close(fig)
                print(f"Gráfico de proyección de Vida Dinámica guardado en: {plot_path}")
            except Exception as e:
                print(f"Error al generar el gráfico de proyección DINÁMICA: {e}")
        else:
            print("Advertencia: No se pudo generar la curva de proyección dinámica para el segmento promedio (ESALs Anuales o K Dinámico es cero).")
    return df
# ====================================================================
# === COMPARATIVOS, UNIFICACIÓN PCI DE FALLA Y MODELO PROPUESTO ======
# ====================================================================

def comparar_modelos(df_vida_util, out_folder):
    out_folder = Path(out_folder)
    out_folder.mkdir(parents=True, exist_ok=True)
    plots_folder = out_folder / "plots_comparativo"
    plots_folder.mkdir(parents=True, exist_ok=True)
    df = df_vida_util.copy()
    # columnas esperadas: Unidad, Vida_Dinamica_FClima_FGeot_Anios, Vida_Util_HDM4_Anios
    expected_cols = ['Unidad', 'Vida_Dinamica_FClima_FGeot_Anios', 'Vida_Util_HDM4_Anios']
    for col in expected_cols:
        if col not in df.columns:
            raise ValueError(f"Falta columna requerida en df_vida_util: {col}")
    df_comp = df[['Unidad', 'Vida_Dinamica_FClima_FGeot_Anios', 'Vida_Util_HDM4_Anios']].copy()
    df_comp['Vida_Dinamica_FClima_FGeot_Anios'] = pd.to_numeric(df_comp['Vida_Dinamica_FClima_FGeot_Anios'], errors='coerce')
    df_comp['Vida_Util_HDM4_Anios'] = pd.to_numeric(df_comp['Vida_Util_HDM4_Anios'], errors='coerce')
    df_comp['Diff_Signed_Dinamica_minus_HDM4'] = df_comp['Vida_Dinamica_FClima_FGeot_Anios'] - df_comp['Vida_Util_HDM4_Anios']
    df_comp['Diff_Abs'] = df_comp['Diff_Signed_Dinamica_minus_HDM4'].abs()
    df_comp['Pct_Error_vs_HDM4'] = np.where(
        (df_comp['Vida_Util_HDM4_Anios'].isna()) | (df_comp['Vida_Util_HDM4_Anios'] == 0),
        np.nan,
        100.0 * df_comp['Diff_Signed_Dinamica_minus_HDM4'] / df_comp['Vida_Util_HDM4_Anios']
    )
    try:
        df_comp['__sort_key'] = df_comp['Unidad'].str.replace('+', '.', regex=False).str.extract(r'(\d+\.?\d*)').astype(float)
        df_comp = df_comp.sort_values('__sort_key').drop(columns='__sort_key')
    except Exception:
        df_comp = df_comp.sort_values('Unidad')
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    excel_path = out_folder / f"Comparativo_HDM4_VidaDinamica_{timestamp}.xlsx"
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df_comp.to_excel(writer, sheet_name='Comparativo', index=False)
        stats = {
            'Metric': [
                'N_segments', 'Mean_VidaDinamica_years', 'Mean_HDM4_years',
                'Mean_AbsDiff_years', 'Mean_PctError_vs_HDM4'
            ],
            'Value': [
                int(df_comp.shape[0]),
                float(df_comp['Vida_Dinamica_FClima_FGeot_Anios'].mean(skipna=True)),
                float(df_comp['Vida_Util_HDM4_Anios'].mean(skipna=True)),
                float(df_comp['Diff_Abs'].mean(skipna=True)),
                float(df_comp['Pct_Error_vs_HDM4'].mean(skipna=True))
            ]
        }
        pd.DataFrame(stats).to_excel(writer, sheet_name='Resumen_Estadistico', index=False)
    try:
        df_plot = df_comp.set_index('Unidad')[[ 'Vida_Dinamica_FClima_FGeot_Anios', 'Vida_Util_HDM4_Anios']].copy()
        fig, ax = plt.subplots(figsize=(12,6))
        df_plot.plot(kind='bar', ax=ax, width=0.8)
        ax.set_title('Comparativo Vida Dinámica (FClima-FGeot) vs HDM-4 Tradicional')
        ax.set_xlabel('Unidad'); ax.set_ylabel('Vida útil restante (años)')
        plt.xticks(rotation=45, ha='right'); plt.tight_layout()
        bar_path = plots_folder / f'Comparativo_VidaDinamica_vs_HDM4_Barras_{timestamp}.png'
        fig.savefig(bar_path, dpi=300); plt.close(fig)
    except Exception as e:
        bar_path = None
        print(f"Error creando gráfico de barras comparativo: {e}")
    try:
        fig, ax = plt.subplots(figsize=(12,5))
        x = range(len(df_comp))
        ax.plot(x, df_comp['Pct_Error_vs_HDM4'], marker='o', linestyle='-', label='% Error (Vida Dinámica vs HDM4)')
        ax.axhline(0, color='gray', linestyle='--', linewidth=0.8)
        ax.set_xticks(x); ax.set_xticklabels(df_comp['Unidad'], rotation=45, ha='right')
        ax.set_ylabel('% Error relativo a HDM-4')
        ax.set_title('Error porcentual: Vida Dinámica vs HDM-4')
        plt.tight_layout()
        pct_path = plots_folder / f'PctError_VidaDinamica_vs_HDM4_{timestamp}.png'
        fig.savefig(pct_path, dpi=300); plt.close(fig)
    except Exception as e:
        pct_path = None
        print(f"Error creando gráfico de % error: {e}")
    dict_paths = {'excel': excel_path, 'bar_plot': bar_path, 'pct_plot': pct_path}
    print(f"✅ Comparación generada. Excel: {excel_path} | Plots: {plots_folder}")
    return df_comp, dict_paths

def unificar_pci_falla_y_comparar(df_vida_util, pci_falla_hdm4, out_folder, pci_falla_original=10.0):
    out_folder = Path(out_folder)
    out_folder.mkdir(parents=True, exist_ok=True)
    plots_folder = out_folder / "plots_unificacion"
    plots_folder.mkdir(parents=True, exist_ok=True)
    df = df_vida_util.copy()
    for col in ['Unidad', 'PCI', 'ESALs_Anuales', 'Vida_Dinamica_FClima_FGeot_Anios', 'Vida_Util_HDM4_Anios']:
        if col not in df.columns:
            raise ValueError(f"Falta columna requerida: {col}")
    def recalcular_vida_dinamica(df0, pci_falla_val):
        d = df0[['Unidad', 'PCI', 'ESALs_Anuales', 'F_Dinamico']].copy()
        d['F_Dinamico'] = df0['F_Dinamico'] if 'F_Dinamico' in df0.columns else 1.0
        K_estandar_local = K_ESTANDAR
        d['K_Dinamico'] = K_estandar_local * (1.0 / d['F_Dinamico'])
        d['PCI_Relacion'] = d['PCI'] / pci_falla_val
        d.loc[d['PCI_Relacion'] <= 1.0, 'N_Falla'] = 0.0
        mask = d['PCI_Relacion'] > 1.0
        d.loc[mask, 'N_Falla'] = (1.0 / d.loc[mask, 'K_Dinamico']) * np.log(d.loc[mask, 'PCI_Relacion'])
        d['Vida_Util'] = d['N_Falla'] / d['ESALs_Anuales'].replace(0, np.nan)
        return d[['Unidad', 'Vida_Util']]
    df_unif = recalcular_vida_dinamica(df, pci_falla_hdm4)
    df_unif.rename(columns={'Vida_Util': 'Vida_Util_Dinamica_Unificado'}, inplace=True)
    df_cmp = df.merge(df_unif, on='Unidad', how='left')
    pci_sens = [pci_falla_original, 15.0, pci_falla_hdm4]
    sens_dict = {}
    for val in pci_sens:
        df_temp = recalcular_vida_dinamica(df, val)
        sens_dict[val] = df_temp.rename(columns={'Vida_Util': f'Vida_Dinamica_PCI_{val}'})
    df_sens = df[['Unidad']].copy()
    for val in pci_sens:
        df_sens = df_sens.merge(sens_dict[val], on='Unidad', how='left')
    try:
        fig, ax = plt.subplots(figsize=(14, 6))
        x = np.arange(len(df_cmp))
        width = 0.28
        ax.bar(x - width, df_cmp['Vida_Dinamica_FClima_FGeot_Anios'], width, label='Dinamica (Original)', alpha=0.9)
        ax.bar(x, df_cmp['Vida_Util_Dinamica_Unificado'], width, label=f'Dinamica (PCI_FALLA = {pci_falla_hdm4:.1f})', alpha=0.9)
        ax.bar(x + width, df_cmp['Vida_Util_HDM4_Anios'], width, label='HDM-4 (Tradicional)', alpha=0.9)
        ax.set_xticks(x); ax.set_xticklabels(df_cmp['Unidad'], rotation=60); ax.set_ylabel("Vida Útil Restante (Años)")
        ax.set_title("Unificación de PCI de Falla: Vida Dinámica vs HDM-4")
        ax.legend(); ax.grid(True, linestyle='--', alpha=0.4); plt.tight_layout()
        plot_path = plots_folder / "Comparativo_Unificado_Dinamica.png"
        plt.savefig(plot_path, dpi=300); plt.close(fig)
    except Exception as e:
        plot_path = None
        print("Error generando gráfico unificado:", e)
    excel_path = out_folder / "Unificacion_PCI_Falla_Dinamica.xlsx"
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df_cmp.to_excel(writer, sheet_name="Comparacion_Unificada", index=False)
        df_sens.to_excel(writer, sheet_name="Sensibilidad_PCI_FALLA", index=False)
    print("✔️ Unificación completada")
    print(f"Excel guardado en: {excel_path}")
    print(f"Gráfico guardado en: {plot_path}")
    return df_cmp, df_sens, excel_path, plot_path
# ====================================================================
# === MODELO PROPUESTO, PIPELINE PRINCIPAL y SALIDAS ==================
# ====================================================================

# MODELO PROPUESTO
def train_and_predict_proposed_model(df_vida_util, base_out_folder, use_target='N_Falla_HDM4'):
    out_folder = Path(base_out_folder)
    out_folder.mkdir(parents=True, exist_ok=True)
    plots_folder = out_folder / "proposed_model"
    plots_folder.mkdir(exist_ok=True)
    df = df_vida_util.copy()
    if 'N_Falla_HDM4' not in df.columns:
        if 'Vida_Util_HDM4_Anios' in df.columns and 'ESALs_Anuales' in df.columns:
            df['N_Falla_HDM4'] = df['Vida_Util_HDM4_Anios'] * df['ESALs_Anuales']
        else:
            raise ValueError("No existe target N_Falla_HDM4 ni Vida_Util_HDM4_Anios/ESALs_Anuales para construirlo.")
    needed_cols = ['Unidad','PCI','IRI','IMDABase_Proyectado','ESALs_Anuales','FClima','FGeot','N_Falla_HDM4']
    for c in needed_cols:
        if c not in df.columns:
            raise ValueError(f"Columna requerida faltante para entrenamiento: {c}")
    df_train = df.dropna(subset=needed_cols).copy()
    if df_train.empty:
        raise ValueError("No hay datos suficientes para entrenar el modelo propuesto.")
    df_train['log_N_falla'] = np.log(df_train['N_Falla_HDM4'].replace(0, np.nan)).replace(-np.inf, np.nan)
    df_train = df_train.dropna(subset=['log_N_falla'])
    df_train['log_IMDA'] = np.log1p(df_train['IMDABase_Proyectado'])
    df_train['PCI_norm'] = df_train['PCI'] / 100.0
    iri_min, iri_max = df_train['IRI'].min(), df_train['IRI'].max()
    if iri_max - iri_min > 0:
        df_train['IRI_norm'] = (df_train['IRI'] - iri_min) / (iri_max - iri_min)
    else:
        df_train['IRI_norm'] = df_train['IRI'] * 0.0
    X = df_train[['PCI_norm','IRI_norm','log_IMDA','FClima','FGeot']].copy()
    X = sm.add_constant(X)
    y = df_train['log_N_falla']
    model = sm.OLS(y, X).fit()
    print("\n=== Resumen modelo propuesto (log N_falla) ===")
    print(model.summary())
    df_pred = df.copy()
    df_pred['log_IMDA'] = np.log1p(df_pred['IMDABase_Proyectado'].fillna(0.0))
    df_pred['PCI_norm'] = df_pred['PCI'] / 100.0
    if iri_max - iri_min > 0:
        df_pred['IRI_norm'] = (df_pred['IRI'] - iri_min) / (iri_max - iri_min)
    else:
        df_pred['IRI_norm'] = df_pred['IRI'] * 0.0
    X_pred = df_pred[['PCI_norm','IRI_norm','log_IMDA','FClima','FGeot']].copy()
    X_pred = sm.add_constant(X_pred.fillna(0.0))
    df_pred['log_N_falla_pred'] = model.predict(X_pred)
    df_pred['N_falla_pred'] = np.exp(df_pred['log_N_falla_pred']).replace([np.inf, -np.inf], np.nan).fillna(0.0)
    df_pred['Vida_Propuesto_Anios'] = df_pred['N_falla_pred'] / df_pred['ESALs_Anuales'].replace(0, np.nan)
    df_pred['Vida_Propuesto_Anios'] = df_pred['Vida_Propuesto_Anios'].clip(lower=0.0).fillna(0.0)
    metrics = {}
    if 'Vida_Util_HDM4_Anios' in df_pred.columns:
        y_true = df_pred['Vida_Util_HDM4_Anios'].to_numpy()
        y_hat = df_pred['Vida_Propuesto_Anios'].to_numpy()
        mask = ~np.isnan(y_true) & ~np.isnan(y_hat)
        if mask.sum() > 0:
            mae = mean_absolute_error(y_true[mask], y_hat[mask])
            rmse = math.sqrt(mean_squared_error(y_true[mask], y_hat[mask]))
            denom = np.where(np.abs(y_true[mask])<1e-6, np.nan, y_true[mask])
            mape = np.nanmean(np.abs((y_hat[mask]-y_true[mask]) / denom)) * 100.0
            r2 = np.corrcoef(y_true[mask], y_hat[mask])[0,1]**2 if mask.sum()>1 else np.nan
            metrics = {'MAE':mae,'RMSE':rmse,'MAPE_%':mape,'R2_approx':r2,'N_samples':int(mask.sum())}
            metrics_df = pd.DataFrame([metrics])
            metrics_df.to_excel(out_folder / 'ProposedModel_Metrics.xlsx', index=False)
        else:
            print("No hay pares (Vida_HDM4, Vida_Propuesto) válidos para cálculo de métricas.")
    else:
        print("No se encontró Vida_Util_HDM4_Anios para evaluar la propuesta frente a HDM-4.")
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    excel_out = out_folder / f"ProposedModel_Predictions_{timestamp}.xlsx"
    df_pred.to_excel(excel_out, index=False, engine='openpyxl')
    try:
        fig, ax = plt.subplots(figsize=(14,6))
        df_plot = df_pred[['Unidad','Vida_Propuesto_Anios','Vida_Util_HDM4_Anios']].copy()
        df_plot = df_plot.set_index('Unidad').sort_index()
        df_plot.plot(kind='bar', ax=ax, width=0.8)
        ax.set_ylabel('Vida (años)'); ax.set_title('Comparativo: Vida_Propuesto vs Vida_HDM4'); plt.xticks(rotation=45, ha='right'); plt.tight_layout()
        bar_path = plots_folder / f'Proposed_vs_HDM4_{timestamp}.png'
        fig.savefig(bar_path, dpi=300); plt.close(fig)
    except Exception as e:
        bar_path = None
        print("Error generando gráfico comparativo propuesto vs HDM4:", e)
    try:
        if 'Vida_Util_HDM4_Anios' in df_pred.columns:
            mask = (~df_pred['Vida_Util_HDM4_Anios'].isna()) & (~df_pred['Vida_Propuesto_Anios'].isna())
            fig, ax = plt.subplots(figsize=(8,8))
            ax.scatter(df_pred.loc[mask,'Vida_Util_HDM4_Anios'], df_pred.loc[mask,'Vida_Propuesto_Anios'], alpha=0.7)
            maxv = max(df_pred.loc[mask,['Vida_Util_HDM4_Anios','Vida_Propuesto_Anios']].max().max(), 1)
            ax.plot([0,maxv],[0,maxv], color='red', linestyle='--')
            ax.set_xlabel('Vida HDM-4 (años)'); ax.set_ylabel('Vida Propuesto (años)'); ax.set_title('Scatter: Vida Propuesto vs Vida HDM-4')
            plt.tight_layout()
            scatter_path = plots_folder / f'Proposed_vs_HDM4_scatter_{timestamp}.png'
            fig.savefig(scatter_path, dpi=300); plt.close(fig)
        else:
            scatter_path = None
    except Exception as e:
        scatter_path = None
        print("Error generando scatter plot:", e)
    paths = {'excel': excel_out, 'bar_plot': bar_path, 'scatter': scatter_path, 'metrics': metrics}
    print("Proposed model predictions saved to:", excel_out)
    print("Metrics:", metrics)
    return df_pred, paths

# ============================
# === PIPELINE PRINCIPAL ====
# ============================
def main_pipeline():
    root = tk.Tk(); root.withdraw()
    messagebox.showinfo('Inicio', 'Selecciona la carpeta de destino para guardar todos los resultados del Pipeline Maestro.')
    base_out_folder = filedialog.askdirectory(title="Seleccionar Carpeta de Destino")
    if not base_out_folder:
        messagebox.showwarning('Advertencia', 'Carpeta de destino no seleccionada. Proceso cancelado.')
        return
    final_summary = {}
    messagebox.showinfo('PCI', 'Selecciona el archivo Excel con los datos de fallas para el cálculo del PCI (Falla, Severidad, Cantidad/Área).')
    pci_file_path = filedialog.askopenfilename(title='Seleccionar Archivo de Datos PCI', filetypes=[('Excel files','*.xls *.xlsx')])
    pci_results = None
    if pci_file_path:
        pci_results = calcular_pci_astm(pci_file_path, out_folder=Path(base_out_folder)/"PCI_Results", ask_folder_if_none=False)
        pci_avg = pci_results['df_resumen']['PCI'].mean() if not pci_results['df_resumen'].empty else np.nan
        final_summary['PCI_Promedio_Global'] = f"{pci_avg:.4f}"
        print(f"\nResultado PCI Promedio: {pci_avg:.4f}")
    df_iri_summary = process_iri_data_and_plot(base_out_folder)
    iri_avg = df_iri_summary['IRI_mean_all'].mean() if df_iri_summary is not None and not df_iri_summary.empty else np.nan
    final_summary['IRI_Promedio_Global'] = f"{iri_avg:.4f}"
    df_merged = None
    model = None
    df_corr = None
    if pci_results and df_iri_summary is not None and not df_iri_summary.empty:
        df_merged = merge_pci_iri(pci_results['df_resumen'], df_iri_summary, base_out_folder)
        model, df_corr = perform_pci_iri_correlation(df_merged, base_out_folder)
        if model:
            final_summary['R2_PCI_vs_IRI'] = f"{model.rsquared:.4f}"
            final_summary['Ecuacion_Regresion'] = f"PCI = {model.params['IRI']:.3f} * IRI + {model.params['const']:.3f}"
    else:
        print("\nSkipping PCI/IRI Merge and Correlation due to missing data.")
        if pci_results and df_iri_summary is not None:
             df_merged = merge_pci_iri(pci_results['df_resumen'], df_iri_summary, base_out_folder)
    FClima_value = calculate_fclima(base_out_folder)
    final_summary['FClima_Final'] = f"{FClima_value:.4f}"
    FGeot_value = get_fgeot_from_aashto()
    final_summary['FGeot_Final'] = f"{FGeot_value:.4f}"
    df_esals_full = None
    esals_params = None
    if pci_results:
        df_esals_full, esals_avg, esals_params = calculate_esals_interpolado(pci_results['df_resumen'], base_out_folder)
        formatted_esals_params = {}
        for k, v in esals_params.items():
            if isinstance(v, float) or isinstance(v, int):
                formatted_esals_params[k] = f"{v:.4f}" if v < 1000 else f"{v:,.2f}"
            else:
                formatted_esals_params[k] = str(v)
        final_summary.update(formatted_esals_params)
    else:
        print("\nSkipping ESALs calculation as PCI data (segment IDs) is missing.")
    df_vida_util = None
    if df_esals_full is not None and df_merged is not None:
        df_model_input = df_esals_full.merge(df_merged[['Unidad', 'IRI']], on='Unidad', how='left')
        # IGC estático + Vida dinámica (FClima x FGeot)
        df_vida_util_calibrado = calculate_igc_and_project_pci(df_model_input, FClima_value, FGeot_value, base_out_folder)
        # HDM-4 vida tradicional
        df_vida_util_hdm4, pci_falla_hdm4_est = project_life_hdm4_iri(df_model_input, model)
        # Merge
        df_vida_util = df_vida_util_calibrado.merge(df_vida_util_hdm4, on='Unidad', how='left')
        igc_avg = df_vida_util['IGC'].mean()
        vida_util_dinamica_avg = df_vida_util['Vida_Dinamica_FClima_FGeot_Anios'].mean()
        vida_util_hdm4_avg = df_vida_util['Vida_Util_HDM4_Anios'].mean()
        final_summary['IGC_Promedio'] = f"{igc_avg:.4f}"
        final_summary['Clasificacion_IGC_Promedio'] = df_vida_util['Clasificacion_IGC'].mode().iloc[0] if 'Clasificacion_IGC' in df_vida_util.columns else ''
        final_summary['Vida_Util_Promedio_Anios_Dinamica'] = f"{vida_util_dinamica_avg:.2f}"
        final_summary['Vida_Util_Promedio_Anios_HDM4'] = f"{vida_util_hdm4_avg:.2f}"
        final_summary['HDM4_IRI_Terminal'] = f"{IRI_TERMINAL_HDM4} m/km"
        final_summary['HDM4_K_Progression'] = f"{K_HDM4_IRI:.2e} 1/ESALs"
        final_summary['HDM4_PCI_Falla_Estimado'] = f"{pci_falla_hdm4_est:.2f}"
        try:
            plots_folder = Path(base_out_folder) / "IGC_Proyeccion_Results"
            plots_folder.mkdir(exist_ok=True)
            df_plot = df_vida_util[['Unidad', 'Vida_Dinamica_FClima_FGeot_Anios', 'Vida_Util_HDM4_Anios']].copy()
            df_plot = df_plot.rename(columns={'Vida_Dinamica_FClima_FGeot_Anios': 'Vida Dinámica (FClima-FGeot)', 'Vida_Util_HDM4_Anios': 'HDM-4 (Tradicional)'})
            pivot_df = df_plot.set_index('Unidad')
            fig, ax = plt.subplots(figsize=(14, 6))
            pivot_df.plot(kind='bar', ax=ax, width=0.8)
            ax.set_title('G6: Comparativo de Vida Dinámica (FClima-FGeot) vs HDM-4 Tradicional')
            ax.set_xlabel('Segmento (Unidad)'); ax.set_ylabel('Vida Útil Restante (Años)')
            plt.xticks(rotation=45, ha='right'); plt.tight_layout()
            fig.savefig(plots_folder / 'G6_Comparativo_VidaDinamica_vs_HDM4.png', dpi=300); plt.close(fig)
        except Exception as e:
            print(f"Error al generar G6: {e}")
        try:
            comparisons_folder = Path(base_out_folder) / "Comparisons"
            df_comp, comp_paths = comparar_modelos(df_vida_util, comparisons_folder)
            final_summary['Comparativo_Excel'] = str(comp_paths['excel'])
            final_summary['Comparativo_Barras'] = str(comp_paths['bar_plot']) if comp_paths.get('bar_plot') else ''
            final_summary['Comparativo_PctPlot'] = str(comp_paths['pct_plot']) if comp_paths.get('pct_plot') else ''
        except Exception as e:
            print(f"Error al generar comparativo Vida Dinámica vs HDM-4: {e}")
        try:
            unif_folder = Path(base_out_folder) / "Unificacion_PCI"
            df_unif_cmp, df_sens, unif_excel, unif_plot = unificar_pci_falla_y_comparar(df_vida_util, pci_falla_hdm4=pci_falla_hdm4_est, out_folder=unif_folder, pci_falla_original=PCI_FALLA)
            final_summary['Unif_PCI_Excel'] = str(unif_excel)
            final_summary['Unif_PCI_Plot'] = str(unif_plot)
        except Exception as e:
            print("ERROR en la unificación de PCI de falla:", e)
        # Modelo propuesto
        try:
            model_folder = Path(base_out_folder) / "ProposedModel"
            df_pred, pred_paths = train_and_predict_proposed_model(df_vida_util, model_folder)
            final_summary['ProposedModel_Excel'] = str(pred_paths.get('excel',''))
            final_summary['ProposedModel_Bar'] = str(pred_paths.get('bar_plot',''))
            final_summary['ProposedModel_Scatter'] = str(pred_paths.get('scatter',''))
            final_summary['ProposedModel_Metrics'] = str(pred_paths.get('metrics',''))
        except Exception as e:
            print("Error entrenando o ejecutando el modelo propuesto:", e)
    else:
        print("\nSkipping model steps due to missing ESALs or merged data.")
    print("\n--- PASO FINAL: Exportando Resultados Consolidados ---")
    results_folder = Path(base_out_folder) / "Pipeline_Consolidado"
    os.makedirs(results_folder, exist_ok=True)
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = results_folder / f"Resultados_Pipeline_Maestro_CONSOLIDADO_v2_{timestamp}.xlsx"
    df_final_results = pd.DataFrame(list(final_summary.items()), columns=['Métrica', 'Valor'])
    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df_final_results.to_excel(writer, sheet_name='0_Resumen_Metodologia', index=False)
            if df_vida_util is not None and not df_vida_util.empty:
                cols_export = ['Unidad', 'PCI', 'IRI', 'IMDABase_Proyectado', 'ESALs_Totales_Diseno', 'ESALs_Anuales', 'IGC', 'Clasificacion_IGC', 'Vida_Dinamica_FClima_FGeot_Anios', 'Vida_Util_HDM4_Anios']
                df_vida_util[cols_export].to_excel(writer, sheet_name='1_Vida_Dinamica_vs_HDM4', index=False)
        print(f"✅ Proceso Finalizado. Resultados consolidados guardados en: {output_path}")
        messagebox.showinfo('Pipeline Finalizado', f"El proceso ha finalizado. Resultados consolidados guardados en:\n{output_path}")
    except Exception as e:
        messagebox.showerror('Error Crítico', f"Fallo durante la exportación final: {e}")
        print(f"Fallo durante la exportación final: {e}")

if __name__ == '__main__':
    try:
        main_pipeline()
    except Exception as e:
        print(f"Ocurrió un error inesperado en el Pipeline Principal: {e}")
