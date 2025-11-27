# -*- coding: utf-8 -*-
"""
PIPELINE MAESTRO COMPLETO (MTC v5.2 - COMPARATIVO HDM-4): 
INTEGRACIÓN DE PARÁMETROS, CÁLCULO DE IGC Y MODELADO DE VIDA ÚTIL (HDM-4 CALIBRADO Y TRADICIONAL)
"""

import pandas as pd
import numpy as np
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import matplotlib.pyplot as plt
import os
import datetime
from scipy.stats import pearsonr
import statsmodels.api as sm
import openpyxl 
import re 
from openpyxl.utils.dataframe import dataframe_to_rows

# ====================================================================
# === 1. CONSTANTES Y FUNCIONES AUXILIARES PCI / IRI / CORRELACIÓN ===
# ====================================================================

# Constantes PCI (Curvas de Deducibilidad - DV y CDV) - Se mantienen sin cambios
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
# === 2. FUNCIONES DE CÁLCULO Y GRAFICACIÓN PCI (ASTM D6433)       ===
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

    # ---------------------------
    # Preparar carpeta de salida y guardar archivos
    # ---------------------------
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
        
    # ---------------------------
    # Gráficos PCI 
    # ---------------------------
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
            if p>=70: return 'Satisfactorio (70-84)'
            if p>=55: return 'Regular (55-69)'
            if p>=40: return 'Malo (40-54)'
            if p>=25: return 'Serio (25-39)'
            return 'Grave (0-24)'
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

# ====================================================================
# === 3. FUNCIONES DE CÁLCULO Y GRAFICACIÓN IRI                    ===
# ====================================================================

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
    
    # ⭐ CORRECCIÓN CLAVE: Asegurar que el Segment ID es un string antes de agrupar y fusionar
    df[seg_col] = df[seg_col].astype(str).str.strip() 
    
    agg_dict = {iri_col: 'mean'}
    if vel_col:
        df[vel_col] = pd.to_numeric(df[vel_col], errors='coerce')
        agg_dict[vel_col] = 'mean'
    
    agg = df.groupby(seg_col).agg(agg_dict).reset_index()

    # ⭐ CORRECCIÓN CLAVE: Renombrar explícitamente a 'Unidad' para el merge
    # agg.columns[0] es la columna resultante de la agrupación (seg_col)
    agg = agg.rename(columns={agg.columns[0]: "Unidad"}) 

    # Renombrar el resto de columnas
    new_cols_iri = ["IRI_m_per_km"] 
    new_cols_vel = ([ "Velocidad_mean_kmh"] if vel_col and len(agg.columns)>2 else [])

    # Creamos un diccionario para renombrar las columnas agrupadas
    rename_dict = {}
    if iri_col in agg.columns:
        rename_dict[iri_col] = new_cols_iri[0]
    if vel_col and len(new_cols_vel) > 0 and vel_col in agg.columns:
        rename_dict[vel_col] = new_cols_vel[0]
        
    agg = agg.rename(columns=rename_dict)
    
    # Doble chequeo de tipo en la salida
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
            # Ajustar nombres de columnas para el merge múltiple
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

    # Merge de todos los recorridos
    merged = iri_agg_list[0].copy() 
    for i in range(1, len(iri_agg_list)):
        # Aseguramos que el merge se hace sobre strings
        iri_agg_list[i]['Unidad'] = iri_agg_list[i]['Unidad'].astype(str)
        merged = merged.merge(iri_agg_list[i], on="Unidad", how="outer")

    # Cálculo del promedio y desviación estándar
    cols_iri = [f"IRI_m_per_km_r{i+1}" for i in range(len(iri_agg_list))]
    cols_vel = [f"Velocidad_r{i+1}" for i in range(len(iri_agg_list)) if f"Velocidad_r{i+1}" in merged.columns]
    
    merged["IRI_mean_all"] = merged[cols_iri].mean(axis=1, skipna=True)
    merged["IRI_std_all"] = merged[cols_iri].std(axis=1, skipna=True)
    merged["n_runs"] = merged[cols_iri].notna().sum(axis=1)

    # Impresión y estadísticas
    iri_recorrido_promedios = {f"Recorrido {i+1}": merged[col].mean(skipna=True) for i, col in enumerate(cols_iri)}
    iri_promedio_total = merged["IRI_mean_all"].mean()
    
    print("\n--- Resultados de Recorridos IRI ---")
    for rec, iri_mean in iri_recorrido_promedios.items():
        print(f"{rec} Promedio: {iri_mean:.3f} m/km")
    print(f"IRI Promedio Total (General): {iri_promedio_total:.3f} m/km")
    
    output_path = out_dir / f"Resultados_IRI_{len(iri_agg_list)}recorridos_merged.xlsx"
    merged.to_excel(output_path, index=False)
    print(f"Resultados IRI guardados en: {output_path}")

    # ---------------------------
    # Gráficos IRI (G1, G2, G3, G4) - Se mantienen sin cambios
    # ---------------------------
    
    merged['Unidad_str'] = merged['Unidad'].astype(str)

    # G1: IRI comparativo de cada recorrido (Líneas)
    try:
        fig, ax = plt.subplots(figsize=(12, 6))
        for col in cols_iri: ax.plot(merged['Unidad_str'], merged[col], marker='o', linestyle='-', label=col.replace('IRI_m_per_km_', ''))
        ax.plot(merged['Unidad_str'], merged['IRI_mean_all'], linestyle='--', color='black', label='IRI Promedio Total'); ax.set_title('IRI por Segmento: Comparativo de Recorridos y Promedio')
        ax.set_xlabel('Segmento (Unidad)'); ax.set_ylabel('IRI (m/km)'); ax.legend(); plt.xticks(rotation=45, ha='right'); plt.grid(axis='y', linestyle='--'); plt.tight_layout()
        fig.savefig(plots_folder / 'G1_IRI_comparativo_recorridos_lineas.png', dpi=300); plt.close(fig)
    except Exception as e: print(f"Error al generar G1 (IRI comparativo de líneas): {e}")

    # G2: IRI promedio de cada recorrido (Barras)
    try:
        rec_names = list(iri_recorrido_promedios.keys()); rec_values = list(iri_recorrido_promedios.values())
        fig, ax = plt.subplots(figsize=(10, 6))
        bars = ax.bar(rec_names, rec_values, color='skyblue')
        ax.axhline(iri_promedio_total, color='red', linestyle='--', label=f'IRI Promedio General ({iri_promedio_total:.2f})'); ax.set_title('IRI Promedio de Cada Recorrido')
        ax.set_ylabel('IRI Promedio (m/km)'); ax.set_xlabel('Recorrido'); ax.legend()
        for bar in bars: yval = bar.get_height(); ax.text(bar.get_x() + bar.get_width()/2.0, yval, f'{yval:.3f}', va='bottom', ha='center')
        plt.tight_layout()
        fig.savefig(plots_folder / 'G2_IRI_promedio_comparativo_barras.png', dpi=300); plt.close(fig)
    except Exception as e: print(f"Error al generar G2 (IRI promedio comparativo barras): {e}")

    # G3: Velocidad por segmento (Líneas)
    if cols_vel:
        try:
            fig, ax = plt.subplots(figsize=(12, 6))
            for col in cols_vel: ax.plot(merged['Unidad_str'], merged[col], marker='.', linestyle='-', label=col.replace('Velocidad_', ''))
            ax.set_title('Velocidad Promedio por Segmento (Recorridos Comparados)'); ax.set_xlabel('Segmento (Unidad)'); ax.set_ylabel('Velocidad (km/h)'); ax.legend()
            plt.xticks(rotation=45, ha='right'); plt.grid(axis='y', linestyle='--'); plt.tight_layout()
            fig.savefig(plots_folder / 'G3_Velocidad_comparativo_lineas.png', dpi=300); plt.close(fig)
        except Exception as e: print(f"Error al generar G3 (Velocidad comparativo): {e}")

    # G4: Distribución del estado superficial (IRI/Rcdec)
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

    # Asegurar que el retorno también es string
    df_return = merged[['Unidad', 'IRI_mean_all']].copy()
    df_return['Unidad'] = df_return['Unidad'].astype(str)
    return df_return

# ====================================================================
# === 4. FUNCIÓN DE INTEGRACIÓN DE RESULTADOS PCI + IRI            ===
# ====================================================================

def merge_pci_iri(df_pci_resumen, df_iri_summary, base_out_folder):
    """
    Combina los resultados finales de PCI y de IRI y los guarda.
    Asegura que 'Unidad' sea string en ambos para el merge.
    """
    df_iri_summary = df_iri_summary.rename(columns={'IRI_mean_all': 'IRI'})
    
    # Aseguramos que 'Unidad' sea string en ambos DFs (doble chequeo de seguridad)
    df_pci_resumen['Unidad'] = df_pci_resumen['Unidad'].astype(str)
    df_iri_summary['Unidad'] = df_iri_summary['Unidad'].astype(str)
    
    df_merged = df_pci_resumen[['Unidad', 'PCI']].merge(
        df_iri_summary, 
        on='Unidad', 
        how='outer'
    )
    
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    excel_path = Path(base_out_folder) / f"Resultados_Integrados_PCI_IRI_{timestamp}.xlsx"
    
    # ⭐ Corrección: Usar openpyxl para la exportación
    df_merged.to_excel(excel_path, index=False, engine='openpyxl')
    
    print(f"\n✅ INTEGRACIÓN PCI + IRI exitosa. Archivo guardado en: {excel_path}")
    
    return df_merged

# ====================================================================
# === 5. CORRELACIÓN Y VALIDACIÓN ESTADÍSTICA (PCI vs IRI)         ===
# ====================================================================

def perform_pci_iri_correlation(df_merged_data, base_out_folder):
    print("\n--- PASO 4: Iniciando Análisis de Correlación PCI vs IRI ---")
    
    plots_folder = Path(base_out_folder) / "Correlation_Results"
    plots_folder.mkdir(exist_ok=True)
    
    # 1. Preparación de Datos (Solo filas con PCI e IRI válidos)
    df_corr = df_merged_data.dropna(subset=['PCI', 'IRI']).copy()
    
    df_corr['PCI'] = pd.to_numeric(df_corr['PCI'], errors='coerce')
    df_corr['IRI'] = pd.to_numeric(df_corr['IRI'], errors='coerce')
    df_corr = df_corr.dropna(subset=['PCI', 'IRI'])
    
    if len(df_corr) < 2:
        print("Advertencia: Menos de 2 muestras válidas para la correlación. Saliendo del análisis.")
        return None, None

    # 2. Correlación de Pearson (r) y Valor p
    r, p_value = pearsonr(df_corr['IRI'], df_corr['PCI'])
    
    # 3. Regresión Lineal (OLS)
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

    # 4. Gráfico de Dispersión y Regresión
    try:
        fig, ax = plt.subplots(figsize=(10, 7))
        
        ax.scatter(df_corr['IRI'], df_corr['PCI'], color='blue', alpha=0.6, label='Datos Muestreados')
        
        # Línea de regresión
        ax.plot(df_corr['IRI'], model.predict(X), color='red', linewidth=2, 
                  label=f'Regresión Lineal\n$PCI = {slope:.3f} \\cdot IRI + {intercept:.3f}$\n$R^2 = {r_squared:.4f}$')
        
        ax.set_title('Correlación entre PCI (Condición) e IRI (Rugosidad)', fontsize=14)
        ax.set_xlabel('IRI Promedio (m/km)', fontsize=12)
        ax.set_ylabel('PCI (Índice de Condición del Pavimento)', fontsize=12)
        ax.grid(True, linestyle='--', alpha=0.6)
        ax.legend(loc='lower left')
        
        plot_path = plots_folder / 'Correlacion_PCI_vs_IRI_Regresion.png'
        fig.savefig(plot_path, dpi=300)
        plt.close(fig)
        print(f"Gráfico de dispersión guardado en: {plot_path}")
    except Exception as e:
        print(f"Error al generar el gráfico de correlación: {e}")
        
    # 5. Guardar resumen estadístico
    stats_summary = pd.DataFrame({
        'Métrica': ['Pearson_r', 'P_value', 'R_squared', 'Intercept', 'Slope', 'N_samples'],
        'Valor': [r, p_value, r_squared, intercept, slope, len(df_corr)]
    })
    stats_path = plots_folder / 'Resumen_Estadistico_PCI_IRI.xlsx'
    stats_summary.to_excel(stats_path, index=False, engine='openpyxl')
    print(f"Resultados estadísticos guardados en: {stats_path}")

    return model, df_corr # Retorna el modelo y el dataframe limpio para el IGC

# ====================================================================
# === 6. CÁLCULO DE ESALs y TRÁFICO (CON INTERPOLACIÓN)            ===
# ====================================================================

# --- COMPLEMENTO: FASE III: CÁLCULO DEL FACTOR CAMIÓN (FC) ---
def load_and_calculate_fc():
    """ Carga la tabla de clasificación y calcula el Factor Camión (FC). """
    root = tk.Tk(); root.withdraw()
    messagebox.showinfo('Tráfico - Factor Camión', 'Selecciona el archivo Excel con la Clasificación Vehicular, Porcentajes y FEA.')
    
    fc_path = filedialog.askopenfilename(title='Seleccionar Archivo de Factor Camión (FC/FEA)', filetypes=[('Excel files','*.xls *.xlsx')])
    if not fc_path:
        messagebox.showwarning('Advertencia', 'Archivo FC no seleccionado. Usando FC por defecto (0.5627).')
        return 0.5627 # Valor por defecto basado en imagen_d4c0d8.png
    
    try:
        df_fc = pd.read_excel(fc_path)
        df_fc.columns = df_fc.columns.str.strip().str.lower().str.replace(' ', '_')

        # Buscar columnas clave: Porcentaje y FEA
        perc_col = next((c for c in df_fc.columns if 'porcentaje' in c or 'pct' in c), None)
        fea_col = next((c for c in df_fc.columns if 'fea' in c), None)

        if not perc_col or not fea_col:
            raise ValueError("Las columnas 'Porcentaje' y 'FEA' no fueron encontradas.")

        df_fc[perc_col] = pd.to_numeric(df_fc[perc_col], errors='coerce').fillna(0)
        df_fc[fea_col] = pd.to_numeric(df_fc[fea_col], errors='coerce').fillna(0)

        # Calcular el Factor Camión: FC = Suma(Porcentaje * FEA)
        df_fc['EAL'] = df_fc[perc_col] * df_fc[fea_col]
        fc_value = df_fc['EAL'].sum()
        
        print(f"  - Factor Camión (FC) calculado desde el archivo: {fc_value:.4f}")
        return fc_value
    
    except Exception as e:
        messagebox.showerror('Error FC', f"Fallo al calcular el Factor Camión: {e}. Usando FC por defecto (0.5627).")
        return 0.5627 # Fallback


def calculate_esals_interpolado(df_pci_resumen, base_out_folder):
    print("\n--- PASO 5: Iniciando Cálculo de ESALs por Metodología de Interpolación ---")

    # --- FASE I & II: INPUT DE PROYECCIÓN Y ATENUACIÓN ---
    root = tk.Tk(); root.withdraw()
    try:
        imda_max_base = simpledialog.askfloat("FASE I: Atenuación", "1. IMDA MÁXIMO (Punto A) en el Año del Estudio (Ej: 450 veh/día):", initialvalue=450.0)
        if imda_max_base is None: return None, None, None
        atenuacion_pct = simpledialog.askfloat("FASE I: Atenuación", "2. Tasa de Atenuación (%) del tráfico a lo largo del tramo (Ej: 20 para 20%):", initialvalue=20.0)
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
        vida_util_diseno = simpledialog.askinteger("FASE IV: Diseño Final", "7. Ingrese la Vida Útil de Diseño (L) en años para la carga acumulada:", initialvalue=20)
        if vida_util_diseno is None: return None, None, None
        carriles_diseno = simpledialog.askinteger("FASE IV: Diseño Final", "8. Ingrese el Número de Carriles en la Dirección de Diseño:", initialvalue=1)
        if carriles_diseno is None: return None, None, None
    except Exception as e:
        messagebox.showerror("Error de Ingreso", f"Error en el ingreso de datos de tráfico: {e}")
        return None, None, None
    # [FIN DE CÓDIGO DE INPUT]

    # --- PROCESAMIENTO INTERNO: INTERPOLACIÓN LINEAL ---
    df_pci_seg = df_pci_resumen.copy()
    df_pci_seg['Unidad'] = df_pci_seg['Unidad'].astype(str)
    
    if 'Unidad' in df_pci_seg.columns:
        # Asume el formato Km+m (ej: 0+150). Usa regex para extraer el número flotante
        # Primero reemplaza el '+' con '.' y luego extrae los números
        df_pci_seg['K'] = df_pci_seg['Unidad'].str.replace('+', '.', regex=False).str.extract(r'(\d+\.?\d*)').astype(float)
    else:
          messagebox.showerror("Error", "No se encontró columna 'Unidad' o 'Progresiva' para la interpolación del IMDA.")
          return None, None, None
    
    df_pci_seg = df_pci_seg.dropna(subset=['K']).sort_values('K')
    if df_pci_seg.empty:
        messagebox.showerror("Error", "La columna de progresiva está vacía o es inválida después de la limpieza.")
        return None, None, None

    K_min = df_pci_seg['K'].min()
    K_max = df_pci_seg['K'].max()
    longitud = K_max - K_min
    
    if longitud > 0:
        # Fórmula: IMDA_i = IMDA_MIN + (IMDA_MAX - IMDA_MIN) * (K_max - Ki) / (K_max - K_min)
        df_pci_seg['IMDA_Interpolado_2014'] = (imda_min_base + (imda_max_base - imda_min_base) * ((K_max - df_pci_seg['K']) / longitud))
    else:
        df_pci_seg['IMDA_Interpolado_2014'] = imda_max_base
    
    # Proyección al Año Base (FASE II)
    df_pci_seg['IMDABase_Proyectado'] = df_pci_seg['IMDA_Interpolado_2014'] * (1 + r_proy)**anos_proyeccion * fe_factor
    IMDABase_i_promedio = df_pci_seg['IMDABase_Proyectado'].mean() # Promedio para el reporte resumen

    # --- FASE III: CÁLCULO DEL FACTOR CAMIÓN (FC) ---
    FC = load_and_calculate_fc()
    
    # --- FASE IV: CÁLCULO FINAL DEL ESALs TOTAL DE DISEÑO ($E_d$) ---
    
    # Factor de Distribución por Carril (Fd)
    Fd = 1.0 / carriles_diseno 

    # Factor de Crecimiento de Diseño (FCD) - Utiliza la fórmula de interés compuesto discreto
    if r_proy > 0:
        FCD = ((1 + r_proy)**vida_util_diseno - 1) / r_proy
    else:
        FCD = vida_util_diseno

    # Cálculo por segmento: $E_{d,i} = IMDABase_{i} \times 365 \times FCD \times Fd \times FC$
    df_pci_seg['ESALs_Totales_Diseno'] = (
        df_pci_seg['IMDABase_Proyectado'] * 365 * FCD * Fd * FC
    )
    
    # Cálculo de ESALs Anuales para el Pronóstico de Vida Útil
    df_pci_seg['ESALs_Anuales'] = (
        df_pci_seg['IMDABase_Proyectado'] * 365 * Fd * FC
    )

    Esals_Totales_Diseno_promedio = df_pci_seg['ESALs_Totales_Diseno'].mean()
    
    print(f"  - Factor de Crecimiento de Diseño (FCD): {FCD:.4f}")
    print(f"  - Factor de Distribución por Carril (Fd): {Fd:.4f}")
    print(f"  - ESALs Totales de Diseño Promedio: {Esals_Totales_Diseno_promedio:.2f}")
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
# === 7. CÁLCULO DEL FACTOR CLIMÁTICO (FClima) [AJUSTADO MTC]      ===
# ====================================================================

def calculate_fclima(base_out_folder):
    """
    Calcula el Factor Climático (FClima) basado en datos históricos de clima.
    Solicita los coeficientes (ch, cp, ct, ct0) al usuario para mayor precisión.
    Utiliza la metodología de producto (FClima = FH * FT) (Ajuste MTC).
    """
    print("\n--- PASO 6: Iniciando Cálculo del Factor Climático (FClima) (Ajuste MTC) ---")
    
    root = tk.Tk(); root.withdraw()
    
    # 1. SOLICITUD DE COEFICIENTES CLIMÁTICOS AL USUARIO
    messagebox.showinfo('Coeficientes Climáticos', 'Ingrese los valores de los coeficientes (ch, cp, ct, ct0) según la zona y el Manual MTC.')
    try:
        # Pide cada coeficiente individualmente con los valores por defecto
        c_h = simpledialog.askfloat("COEFICIENTES", "Coeficiente de Humedad (ch) [Ej: 0.0002 para Sierra]:", initialvalue=0.0002)
        if c_h is None: return 1.0
        c_p = simpledialog.askfloat("COEFICIENTES", "Coeficiente Base de Humedad (cp) [Ej: 1.0]:", initialvalue=1.0)
        if c_p is None: return 1.0
        c_t = simpledialog.askfloat("COEFICIENTES", "Coeficiente de Temperatura (ct) [Ej: 0.020 para Sierra]:", initialvalue=0.020)
        if c_t is None: return 1.0
        c_t0 = simpledialog.askfloat("COEFICIENTES", "Coeficiente Base de Temperatura (ct0) [Ej: 0.8]:", initialvalue=0.8)
        if c_t0 is None: return 1.0
    except Exception as e:
        messagebox.showerror('Error de Ingreso', f"Fallo al ingresar los coeficientes: {e}. Asumiendo FClima = 1.0.")
        return 1.0
    
    # 2. SOLICITUD DEL ARCHIVO DE DATOS DE CLIMATOLOGÍA
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
        
        # CÁLCULO DE PROMEDIOS HISTÓRICOS GLOBALES
        df_clima['Temp_Media_Mensual'] = (df_clima['TempMax'] + df_clima['TempMin']) / 2.0
        T_media_avg = df_clima['Temp_Media_Mensual'].mean() 
        df_anual = df_clima.groupby('Año').agg({'Precip': 'sum'}).reset_index()
        P_anual_avg = df_anual['Precip'].mean() 

        # 3. CÁLCULO USANDO LOS COEFICIENTES INGRESADOS POR EL USUARIO
        
        # Cálculo de Factores Parciales
        F_H = max(0.01, c_p - c_h * P_anual_avg) 
        F_T = max(0.01, c_t0 + c_t * T_media_avg)
        
        # FÓRMULA AJUSTADA (MTC/Severidad): F_Clima = F_H * F_T 
        F_Clima = F_H * F_T
        
        print(f"  - Coeficientes usados: ch={c_h:.4f}, cp={c_p:.2f}, ct={c_t:.4f}, ct0={c_t0:.2f}")
        print(f"  - T. Media Histórica: {T_media_avg:.4f} °C | Precip. Anual Promedio: {P_anual_avg:.4f} mm/año")
        print(f"  - Factor de Humedad (F_H): {F_H:.4f} | Factor de Temperatura (F_T): {F_T:.4f}")
        print(f"  - Factor Climático (FClima) [Ajustado MTC]: {F_Clima:.4f}")
        print("✅ FClima calculado con fórmula de producto y coeficientes personalizados.")

        return F_Clima

    except Exception as e:
        messagebox.showerror('Error en FClima', f"Fallo al calcular FClima: {e}. Asumiendo FClima = 1.0 (Neutral).")
        print(f"Error en FClima: {e}")
        return 1.0 # Retorna 1.0 como valor neutral si hay un fallo
        
# ====================================================================
# === 8. CÁLCULO DEL FACTOR GEOTÉCNICO (FGeot) [AJUSTADO POR TABLA] ===
# ====================================================================

# Tabla 23: Normalización del factor de riesgo geotécnico (FGeot)
FGEOT_TABLE = {
    'A-1-A': 1.00,
    'A-1-B': 0.90,
    'A-2': 0.80,
    'A-3': 0.70,
    'A-4': 0.50,
    'A-5': 0.40,
    'A-6': 0.25,
    'A-7-5': 0.15,
    'A-7-6': 0.10
}

def get_fgeot_from_aashto():
    """
    Solicita la clasificación AASHTO de la subrasante y devuelve
    el Factor Geotécnico (FGeot) normalizado según la Tabla 23 del usuario.
    """
    print("\n--- PASO 7: Iniciando Cálculo del Factor Geotécnico (FGeot) ---")
    
    root = tk.Tk(); root.withdraw()
    
    opciones = list(FGEOT_TABLE.keys())
    
    messagebox.showinfo('Factor Geotécnico', 'Seleccione la Clasificación AASHTO obtenida de la calicata para asignar el Factor Geotécnico (FGeot).')
    
    clase_suelo = simpledialog.askstring(
        "Clasificación AASHTO", 
        f"Ingrese la clasificación AASHTO (Ej: A-1-B, A-4, etc.). Opciones: {', '.join(opciones)}:",
        initialvalue='A-1-B' # Valor por defecto A-1-B
    )

    if not clase_suelo:
        messagebox.showwarning('Advertencia', 'Clasificación AASHTO no ingresada. Asumiendo FGeot = 1.0 (Riesgo Mínimo).')
        return 1.0

    # Normalizar la entrada para que coincida con la tabla (mayúsculas, guiones)
    clase_suelo = clase_suelo.strip().upper().replace('A-1-A','A-1-A').replace('A-1-B','A-1-B').replace('A7-5','A-7-5').replace('A7-6','A-7-6')
    
    fgeot_value = FGEOT_TABLE.get(clase_suelo)
    
    if fgeot_value is not None:
        print(f"  - Clasificación AASHTO ingresada: {clase_suelo}")
        print(f"  - Factor Geotécnico (FGeot) asignado: {fgeot_value:.2f}")
        print("✅ FGeot asignado correctamente según la Tabla 23.")
        return fgeot_value
    else:
        messagebox.showerror('Error de Clasificación', f"Clasificación '{clase_suelo}' no encontrada en la Tabla. Asumiendo FGeot = 1.0 (Riesgo Mínimo).")
        return 1.0

# ====================================================================
# === 9. MODELADO COMPUTACIONAL: IGC Y PROYECCIÓN DE VIDA ÚTIL     ===
# ====================================================================

# Ponderaciones AHP Promedio (Tabla 26)
W_AHP = {
    'PCI': 0.4954,
    'IRI': 0.2417,
    'IMDA': 0.1251,
    'Clima': 0.0876,
    'Suelo': 0.0502
}

# Constantes de Modelado de Vida Útil (Metodología Calibrada IGC)
PCI_FALLA = 10.0              # Umbral de falla crítica/estructural (¡Corregido!)
K_ESTANDAR = 2.0e-6           # Coeficiente de deterioro base (ESALs^-1)

# Constantes de Modelado de Vida Útil (Metodología Tradicional HDM-4 - Roughness Progression)
IRI_TERMINAL_HDM4 = 8.0       # IRI Terminal para Rehabilitación (MTC/HDM-4 estándar)
K_HDM4_IRI = 3.0e-5           # Coeficiente de progresión de rugosidad (1/ESALs). Valor típico para un pavimento flexible con un SN promedio.


def get_pci_category_igc(igc):
    """Clasifica el estado del pavimento según el IGC."""
    # Asume IGC en escala 0-1
    if igc >= 0.85: return 'Excelente'
    if igc >= 0.70: return 'Satisfactorio'
    if igc >= 0.55: return 'Regular'
    if igc >= 0.40: return 'Malo'
    if igc >= 0.25: return 'Serio'
    return 'Grave'


def project_life_hdm4_iri(df_model_input, pci_iri_model):
    """
    Proyecta la vida útil restante usando un modelo simplificado de progresión de IRI
    basado en HDM-4 (solo componente de tráfico) para llegar al IRI Terminal.
    """
    df = df_model_input.copy()
    
    # 1. Parámetros de la correlación PCI-IRI (desde el análisis OLS)
    # Se utiliza el modelo local de la Fase 5 para estimar el PCI de falla HDM-4
    if pci_iri_model and 'IRI' in pci_iri_model.params:
        slope = pci_iri_model.params['IRI']
        intercept = pci_iri_model.params['const']
    else:
        slope, intercept = -10.0, 100.0 

    # 2. Calcular el PCI equivalente al IRI Terminal de HDM-4 (8.0 m/km)
    pci_falla_equivalente = max(0.0, slope * IRI_TERMINAL_HDM4 + intercept)
    
    # 3. Cálculo de ESALs Acumulados hasta la Falla (N_Falla_HDM4)
    # Modelo simplificado HDM-4 Roughness Progression: IRI_t = IRI_0 + K * N
    # N_Falla = (IRI_Terminal - IRI_0) / K_HDM4_IRI
    
    df['Delta_IRI_Falla'] = IRI_TERMINAL_HDM4 - df['IRI']
    
    # Manejar el caso donde el IRI actual ya es mayor o igual al IRI terminal
    df.loc[df['Delta_IRI_Falla'] <= 0, 'N_Falla_HDM4'] = 0.0
    
    # Calcular los ESALs necesarios
    df.loc[df['Delta_IRI_Falla'] > 0, 'N_Falla_HDM4'] = (
        df['Delta_IRI_Falla'] / K_HDM4_IRI
    )

    # 4. CÁLCULO DE VIDA ÚTIL RESTANTE HDM-4 (Años)
    # Vida Útil = N_Falla / ESALs Anuales
    df['Vida_Util_HDM4_Anios'] = df['N_Falla_HDM4'] / df['ESALs_Anuales'].replace(0, np.nan)
    df['Vida_Util_HDM4_Anios'] = df['Vida_Util_HDM4_Anios'].clip(lower=0.0).fillna(0.0) # Asegurar no negativos

    print(f"\n--- Resultados de Proyección HDM-4 (Tradicional) ---")
    print(f"  - IRI Terminal asumido: {IRI_TERMINAL_HDM4} m/km")
    print(f"  - K_HDM4 (Progresión IRI) asumido: {K_HDM4_IRI:.2e} 1/ESALs")
    print(f"  - PCI Equivalente de Falla (Estimado): {pci_falla_equivalente:.2f}")
    print(f"  - Vida Útil HDM-4 Restante Promedio: {df['Vida_Util_HDM4_Anios'].mean():.2f} años")
    print("✅ Proyección HDM-4 completada.")
    
    # Devolver un resumen de los resultados HDM-4 para la tabla de comparación
    return df[['Unidad', 'Vida_Util_HDM4_Anios']], pci_falla_equivalente


def calculate_igc_and_project_pci(df_merged_pci_esals, FClima, FGeot, base_out_folder):
    """
    Fase I: Calcula el IGC (Normalización + AHP).
    Fase II: Proyecta la vida útil restante usando el IGC para calibrar el factor K.
    """
    print("\n--- PASO 8: Iniciando Modelado de IGC y Vida Útil Proyectada (Calibrado) ---")

    df = df_merged_pci_esals.copy()
    
    # 1. NORMALIZACIÓN DE PARÁMETROS (Ni)
    
    # Rango de IRI (m/km) para normalización: [2.0 (óptimo) a 8.0 (falla MTC)]
    IRI_MAX = 8.0
    IRI_MIN = 2.0
    
    # Rango de IMDA (veh/día) para normalización
    IMDA_MAX = df['IMDABase_Proyectado'].max()
    IMDA_MIN = df['IMDABase_Proyectado'].min()

    # Normalización del PCI (Mayor es mejor)
    df['N_PCI'] = df['PCI'] / 100.0

    # Normalización del IRI (Menor es mejor)
    df['N_IRI'] = (IRI_MAX - df['IRI']) / (IRI_MAX - IRI_MIN)
    df['N_IRI'] = df['N_IRI'].clip(0.0, 1.0) # Acotando a [0, 1]

    # Normalización del IMDA (Mayor tráfico = Menor conservación/Más riesgo)
    # Se usa la inversa del riesgo
    if (IMDA_MAX - IMDA_MIN) > 0:
        df['N_IMDA'] = (IMDA_MAX - df['IMDABase_Proyectado']) / (IMDA_MAX - IMDA_MIN)
    else: # Si el IMDA es constante
        df['N_IMDA'] = 1.0 
    df['N_IMDA'] = df['N_IMDA'].clip(0.0, 1.0)

    # N_Clima y N_Geot ya están en escala 0-1 (FClima y FGeot)
    df['N_Clima'] = FClima
    df['N_Suelo'] = FGeot

    # 2. CÁLCULO DEL ÍNDICE GLOBAL DE CONSERVACIÓN (IGC) - FASE I
    df['IGC'] = (
        W_AHP['PCI'] * df['N_PCI'] +
        W_AHP['IRI'] * df['N_IRI'] +
        W_AHP['IMDA'] * df['N_IMDA'] +
        W_AHP['Clima'] * df['N_Clima'] +
        W_AHP['Suelo'] * df['N_Suelo']
    )
    df['Clasificacion_IGC'] = df['IGC'].apply(get_pci_category_igc)
    print(f"  - IGC Promedio del tramo: {df['IGC'].mean():.4f}")
    
    # 3. PROYECCIÓN DE VIDA ÚTIL RESTANTE - FASE II (Modelo PCI Calibrado)
    
    # Factor de Aceleración (Inversa del IGC, para que un IGC bajo acelere el deterioro)
    df['F_Aceleracion'] = 1.0 / df['IGC'].replace(0, 1.0) # Evita división por cero
    
    # Coeficiente K Ajustado (K_Ajustado = K_Estandar * F_Aceleracion)
    df['K_Ajustado'] = K_ESTANDAR * df['F_Aceleracion']

    # ESALs Acumulados hasta la Falla (N_Falla)
    # N_Falla = (1 / K_Ajustado) * ln(PCI_0 / PCI_FALLA)
    # PCI_0 es el PCI medido (columna 'PCI')
    
    # Manejar logaritmo de (PCI_0 / PCI_FALLA) cuando PCI_0 <= PCI_FALLA (ya fallado)
    df['PCI_Relacion'] = df['PCI'] / PCI_FALLA
    
    # Caso 1: Pavimento ya fallado (PCI <= PCI_FALLA)
    df.loc[df['PCI_Relacion'] <= 1.0, 'N_Falla'] = 0.0
    
    # Caso 2: Pavimento con vida restante
    df.loc[df['PCI_Relacion'] > 1.0, 'N_Falla'] = (
        (1.0 / df['K_Ajustado']) * np.log(df['PCI_Relacion'])
    )

    # 4. CÁLCULO DE VIDA ÚTIL RESTANTE (Años)
    # Vida Útil = N_Falla / ESALs Anuales
    df['Vida_Util_IGC_Anios'] = df['N_Falla'] / df['ESALs_Anuales'].replace(0, np.nan)
    df['Vida_Util_IGC_Anios'] = df['Vida_Util_IGC_Anios'].clip(lower=0.0).fillna(0.0)
    
    print(f"  - K Estándar: {K_ESTANDAR} | PCI de Falla: {PCI_FALLA}")
    print(f"  - Vida Útil Restante Promedio (IGC): {df['Vida_Util_IGC_Anios'].mean():.2f} años")
    print("✅ Modelado IGC y Vida Útil Calibrada completado.")

    # 5. GRÁFICOS DE PROYECCIÓN (Ejemplo para el segmento con el IGC promedio)
    
    plots_folder = Path(base_out_folder) / "IGC_Proyeccion_Results"
    plots_folder.mkdir(exist_ok=True)
    
    # Usar el segmento con el PCI más cercano al promedio del tramo
    pci_promedio = df['PCI'].mean()
    idx_ejemplo = (df['PCI'] - pci_promedio).abs().argsort().iloc[0]
    segmento_ejemplo = df.loc[idx_ejemplo]
    
    # Generar la curva de proyección
    pci_0_seg = segmento_ejemplo['PCI']
    k_ajustado_seg = segmento_ejemplo['K_Ajustado']
    esals_anuales_seg = segmento_ejemplo['ESALs_Anuales']
    
    if esals_anuales_seg > 0 and k_ajustado_seg > 0:
        
        # Generar puntos de ESALs (hasta 2x la vida útil calculada)
        max_esals = segmento_ejemplo['N_Falla'] * 1.5
        esals_range = np.linspace(0, max_esals, 100)
        
        # Calcular la caída del PCI
        pci_proyectado = pci_0_seg * np.exp(-k_ajustado_seg * esals_range)
        
        # Calcular los años
        años_proyectados = esals_range / esals_anuales_seg

        try:
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.plot(años_proyectados, pci_proyectado, label=f'PCI Proyectado (K_Ajustado={k_ajustado_seg:.2e})', color='blue')
            ax.axhline(PCI_FALLA, color='red', linestyle='--', label=f'Umbral de Falla (PCI={PCI_FALLA})')
            ax.scatter(segmento_ejemplo['Vida_Util_IGC_Anios'], PCI_FALLA, color='red', marker='o', s=100, zorder=5, 
                       label=f"Vida Útil Restante (IGC): {segmento_ejemplo['Vida_Util_IGC_Anios']:.2f} años")
            
            ax.set_title(f"G5: Proyección de Vida Útil (Segmento: {segmento_ejemplo['Unidad']}) - Método Calibrado", fontsize=14)
            ax.set_xlabel('Años a partir del Estudio Base', fontsize=12)
            ax.set_ylabel('PCI', fontsize=12)
            ax.set_ylim(0, 100)
            ax.grid(True, linestyle='--', alpha=0.6)
            ax.legend(loc='upper right')
            
            plot_path = plots_folder / f"G5_Proyeccion_Vida_Util_Segmento_{segmento_ejemplo['Unidad'].replace('+','-')}.png"
            fig.savefig(plot_path, dpi=300)
            plt.close(fig)
            print(f"Gráfico de proyección de PCI (Calibrado) guardado en: {plot_path}")
        except Exception as e:
            print(f"Error al generar el gráfico de proyección PCI: {e}")
    else:
        print("Advertencia: No se pudo generar la curva de proyección para el segmento promedio (ESALs Anuales o K Ajustado es cero).")
        
    return df

# ====================================================================
# === 10. FUNCIÓN DE PIPELINE MAESTRO Y EXPORTACIÓN [COMPLETADA]   ===
# ====================================================================

def main_pipeline():
    """ Orquesta todo el proceso y consolida los resultados. """
    
    root = tk.Tk(); root.withdraw()
    messagebox.showinfo('Inicio', 'Selecciona la carpeta de destino para guardar todos los resultados del Pipeline Maestro.')
    base_out_folder = filedialog.askdirectory(title="Seleccionar Carpeta de Destino")
    
    if not base_out_folder:
        messagebox.showwarning('Advertencia', 'Carpeta de destino no seleccionada. Proceso cancelado.')
        return
        
    final_summary = {}
    
    # --- PASO 1: PCI (ASTM D6433) ---
    messagebox.showinfo('PCI', 'Selecciona el archivo Excel con los datos de fallas para el cálculo del PCI (Falla, Severidad, Cantidad/Área).')
    pci_file_path = filedialog.askopenfilename(title='Seleccionar Archivo de Datos PCI', filetypes=[('Excel files','*.xls *.xlsx')])
    
    pci_results = None
    if pci_file_path:
        pci_results = calcular_pci_astm(pci_file_path, out_folder=Path(base_out_folder)/"PCI_Results", ask_folder_if_none=False)
        pci_avg = pci_results['df_resumen']['PCI'].mean() if not pci_results['df_resumen'].empty else np.nan
        final_summary['PCI_Promedio_Global'] = f"{pci_avg:.4f}"
        print(f"\nResultado PCI Promedio: {pci_avg:.4f}")
    
    # --- PASO 2: IRI (Multiple Runs Aggregation) ---
    df_iri_summary = process_iri_data_and_plot(base_out_folder)
    iri_avg = df_iri_summary['IRI_mean_all'].mean() if df_iri_summary is not None and not df_iri_summary.empty else np.nan
    final_summary['IRI_Promedio_Global'] = f"{iri_avg:.4f}"
    
    # --- PASO 3: INTEGRACIÓN PCI + IRI y Correlación ---
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
        # Crear un DF dummy para los pasos siguientes si solo falta la correlación
        if pci_results and df_iri_summary is not None:
             df_merged = merge_pci_iri(pci_results['df_resumen'], df_iri_summary, base_out_folder)

    # --- PASO 4: FACTOR CLIMÁTICO (FClima) ---
    FClima_value = calculate_fclima(base_out_folder)
    final_summary['FClima_Final'] = f"{FClima_value:.4f}"

    # --- PASO 5: FACTOR GEOTÉCNICO (FGeot) ---
    FGeot_value = get_fgeot_from_aashto()
    final_summary['FGeot_Final'] = f"{FGeot_value:.4f}"
    
    # --- PASO 6: TRÁFICO (ESALs) ---
    df_esals_full = None
    esals_params = None
    if pci_results:
        df_esals_full, esals_avg, esals_params = calculate_esals_interpolado(pci_results['df_resumen'], base_out_folder)
        # Formatear los resultados de tráfico para el resumen
        formatted_esals_params = {}
        for k, v in esals_params.items():
            if isinstance(v, float) or isinstance(v, int):
                formatted_esals_params[k] = f"{v:.4f}" if v < 1000 else f"{v:,.2f}"
            else:
                formatted_esals_params[k] = str(v)
        final_summary.update(formatted_esals_params)
    else:
        print("\nSkipping ESALs calculation as PCI data (segment IDs) is missing.")

    # --- PASO 7: MODELADO IGC y VIDA ÚTIL (Calibrado) ---
    df_vida_util = None
    if df_esals_full is not None and df_merged is not None:
        # Fusionar los ESALs y los datos PCI/IRI
        df_model_input = df_esals_full.merge(df_merged[['Unidad', 'IRI']], on='Unidad', how='left')
        
        # Aplicar el Modelado IGC (Calibrado)
        df_vida_util_calibrado = calculate_igc_and_project_pci(
            df_model_input, FClima_value, FGeot_value, base_out_folder
        )
        
        # -----------------------------------------------------
        # ⭐ PASO 8: MODELADO HDM-4 TRADICIONAL (IRI Progression) ⭐
        # -----------------------------------------------------
        df_vida_util_hdm4, pci_falla_hdm4_est = project_life_hdm4_iri(df_model_input, model)
        
        # Combinar resultados Calibrado y HDM-4
        df_vida_util = df_vida_util_calibrado.merge(df_vida_util_hdm4, on='Unidad', how='left')
        
        # ⭐ RESULTADOS CLAVE PARA EL RESUMEN
        igc_avg = df_vida_util['IGC'].mean()
        vida_util_igc_avg = df_vida_util['Vida_Util_IGC_Anios'].mean()
        vida_util_hdm4_avg = df_vida_util['Vida_Util_HDM4_Anios'].mean()
        
        final_summary['IGC_Promedio'] = f"{igc_avg:.4f}"
        final_summary['Clasificacion_IGC_Promedio'] = get_pci_category_igc(igc_avg)
        final_summary['Vida_Util_Promedio_Anios_IGC'] = f"{vida_util_igc_avg:.2f}"
        final_summary['Vida_Util_Promedio_Anios_HDM4'] = f"{vida_util_hdm4_avg:.2f}"
        final_summary['HDM4_IRI_Terminal'] = f"{IRI_TERMINAL_HDM4} m/km"
        final_summary['HDM4_K_Progression'] = f"{K_HDM4_IRI:.2e} 1/ESALs"
        final_summary['HDM4_PCI_Falla_Estimado'] = f"{pci_falla_hdm4_est:.2f}"
        
        # Gráfico Comparativo
        try:
            plots_folder = Path(base_out_folder) / "IGC_Proyeccion_Results"
            
            df_plot = df_vida_util[['Unidad', 'Vida_Util_IGC_Anios', 'Vida_Util_HDM4_Anios']].copy()
            df_plot = df_plot.rename(columns={'Vida_Util_IGC_Anios': 'IGC (Calibrado)', 'Vida_Util_HDM4_Anios': 'HDM-4 (Tradicional)'})
            df_plot_melt = df_plot.melt(id_vars='Unidad', var_name='Metodología', value_name='Vida_Util_Anios')

            fig, ax = plt.subplots(figsize=(12, 6))
            pivot_df = df_plot.set_index('Unidad')
            pivot_df.plot(kind='bar', ax=ax, width=0.8)
            
            ax.set_title('G6: Comparativo de Vida Útil Restante (IGC Calibrado vs. HDM-4 Tradicional)', fontsize=14)
            ax.set_xlabel('Segmento (Unidad)', fontsize=12)
            ax.set_ylabel('Vida Útil Restante (Años)', fontsize=12)
            plt.xticks(rotation=45, ha='right')
            ax.grid(axis='y', linestyle='--', alpha=0.6)
            plt.tight_layout()
            
            plot_path = plots_folder / "G6_Comparativo_Vida_Util.png"
            fig.savefig(plot_path, dpi=300)
            plt.close(fig)
            print(f"Gráfico comparativo de vida útil guardado en: {plot_path}")
            
        except Exception as e:
             print(f"Error al generar el gráfico comparativo G6: {e}")
            
    else:
        print("\nSkipping IGC, HDM-4 and Life Prediction due to missing ESALs or PCI/IRI data.")


    # --- 9. EXPORTACIÓN FINAL CONSOLIDADA ---
    print("\n--- PASO FINAL: Exportando Resultados Consolidados ---")
    
    results_folder = Path(base_out_folder) / "Pipeline_Consolidado"
    os.makedirs(results_folder, exist_ok=True)
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = results_folder / f"Resultados_Pipeline_Maestro_CONSOLIDADO_{timestamp}.xlsx"
    
    # Asegurar que el resumen es un DataFrame bien formado para la exportación
    df_final_results = pd.DataFrame(list(final_summary.items()), columns=['Métrica', 'Valor'])

    try:
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Resumen de Métricas
            df_final_results.to_excel(writer, sheet_name='0_Resumen_Metodologia', index=False)
            
            # Resultados por Segmento (IGC, HDM-4 y Vida Útil)
            if df_vida_util is not None and not df_vida_util.empty:
                cols_export = ['Unidad', 'PCI', 'IRI', 'IMDABase_Proyectado', 
                               'ESALs_Totales_Diseno', 'ESALs_Anuales', 'IGC', 
                               'Clasificacion_IGC', 'Vida_Util_IGC_Anios',
                               'Vida_Util_HDM4_Anios'] # Añadida columna HDM-4
                
                df_vida_util[cols_export].to_excel(writer, sheet_name='1_IGC_HDM4_Segmento', index=False)
        
        print(f"✅ Proceso Finalizado. Resultados consolidados guardados en: {output_path}")
        messagebox.showinfo('Pipeline Finalizado', f"El proceso ha finalizado. Resultados consolidados guardados en:\n{output_path}")

    except Exception as e:
        messagebox.showerror('Error Crítico', f"Fallo durante la exportación final: {e}")
        print(f"Fallo durante la exportación final: {e}")


if __name__ == '__main__':
    # Usar try-except para que los errores de Tkinter o Matplotlib no colapsen la consola
    try:
        main_pipeline()
    except Exception as e:
        print(f"Ocurrió un error inesperado en el Pipeline Principal: {e}")
