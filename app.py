import pandas as pd
import os
import re
from datetime import datetime

# Patrón del nombre del archivo para identificar el mes
pattern = r"detenciones-Llenado_V2-(\d{6})\d{2}-\d{8}\.xlsx"

# Diccionarios para acumular horas por mes
monthly_hours = {}
monthly_production_hours = {}
monthly_failures_hours = {}
monthly_scheduled_hours = {}
monthly_maintenance_hours = {}
monthly_unplanned_hours = {}
monthly_micro_stops_hours = {}

def find_duration_column(df, verbose=False):
    if verbose:
        print("Columnas encontradas:", list(df.columns))
    
    # Lista de posibles nombres de columna
    possible_names = [
        "DURACION EN MINUTOS",
        "DURACIÓN EN MINUTOS", 
        "DURACION_EN_MINUTOS",
        "duracion en minutos",
        "Duracion en Minutos"
    ]
    
    # Buscar coincidencia exacta primero
    for col in df.columns:
        col_clean = str(col).strip()
        if col_clean in possible_names:
            return col_clean
    
    # Buscar por contenido (más flexible)
    for col in df.columns:
        col_upper = str(col).upper().strip()
        if "DURACION" in col_upper and "MINUTO" in col_upper:
            return col
    
    # Si no encontramos nada específico, buscar columnas que puedan contener números y tengan nombres relacionados con tiempo
    for col in df.columns:
        col_upper = str(col).upper().strip()
        if any(word in col_upper for word in ["TIEMPO", "MINUTO", "DURACION", "DURATION"]):
            # Verificar si la columna tiene datos numéricos
            try:
                numeric_data = pd.to_numeric(df[col], errors='coerce').dropna()
                if len(numeric_data) > 0:
                    return col
            except:
                continue
    
    return None

# Procesar archivos Excel
for filename in os.listdir("."):
    if filename.endswith('.xlsx') and 'detenciones-Llenado_V2' in filename:
        match = re.match(pattern, filename)
        if match:
            yyyymm = match.group(1)
        else:
            # Si el patrón no coincide exactamente, intentar extraer la fecha del nombre
            date_match = re.search(r'(\d{6})', filename)
            if date_match:
                yyyymm = date_match.group(1)
            else:
                print(f"No se pudo extraer fecha de {filename}")
                continue
        
        try:            
            # Primero intentar leer normalmente
            df = pd.read_excel(filename, sheet_name='detalle')
            
            # Si las columnas son "Unnamed", intentar con diferentes filas como header
            if all(col.startswith('Unnamed:') for col in df.columns):
                for header_row in range(0, min(10, len(df))):
                    try:
                        df_test = pd.read_excel(filename, sheet_name='detalle', header=header_row)
                        duration_col = find_duration_column(df_test, verbose=False)
                        if duration_col:
                            df = df_test
                            break
                    except Exception as e:
                        continue
            
            # Buscar la columna de duración
            duration_col = find_duration_column(df, verbose=False)
            
            if duration_col:
                try:
                    df[duration_col] = pd.to_numeric(df[duration_col], errors='coerce')
                    
                    # Calcular total general
                    total_minutes = df[duration_col].sum()
                    total_hours = total_minutes / 60
                    monthly_hours[yyyymm] = total_hours
                    
                    # Calcular total de producción
                    production_mask = df['CODIGO DETENCION'].str.lower().isin(['produccion', 'producción'])
                    production_data = df.loc[production_mask, duration_col]
                    production_minutes = production_data.sum()
                    production_hours = production_minutes / 60
                    monthly_production_hours[yyyymm] = production_hours
                    
                    # Calcular total de micro paradas
                    micro_stops_mask = df['CODIGO DETENCION'].str.lower().str.contains('micro parada', na=False)
                    micro_stops_data = df.loc[micro_stops_mask, duration_col]
                    micro_stops_minutes = micro_stops_data.sum()
                    micro_stops_hours = micro_stops_minutes / 60
                    monthly_micro_stops_hours[yyyymm] = micro_stops_hours
                    
                    # Calcular total de paradas no planificadas
                    unplanned_base_mask = df['CODIGO DETENCION'].str.lower().str.contains('2. paradas no planificadas', na=False)
                    
                    # Dentro de las no planificadas, identificar fallas y averías
                    failures_mask = unplanned_base_mask & df['CODIGO DETENCION'].str.lower().str.contains('fallas y averias|fallas y averías', na=False)
                    failures_data = df.loc[failures_mask, duration_col]
                    failures_minutes = failures_data.sum()
                    failures_hours = failures_minutes / 60
                    monthly_failures_hours[yyyymm] = failures_hours

                    # Calcular otras paradas no planificadas (excluyendo fallas y averías)
                    unplanned_mask = unplanned_base_mask & ~failures_mask
                    unplanned_data = df.loc[unplanned_mask, duration_col]
                    unplanned_minutes = unplanned_data.sum()
                    unplanned_hours = unplanned_minutes / 60
                    monthly_unplanned_hours[yyyymm] = unplanned_hours
                    
                    # Calcular total de paradas programada
                    scheduled_base_mask = df['CODIGO DETENCION'].str.lower().str.contains('1. paradas programadas', na=False)
                    
                    # Identificar mantención dentro de las programadas
                    maintenance_mask = scheduled_base_mask & df['CODIGO DETENCION'].str.lower().str.contains('mantención|mantencion', na=False)
                    maintenance_data = df.loc[maintenance_mask, duration_col]
                    maintenance_minutes = maintenance_data.sum()
                    maintenance_hours = maintenance_minutes / 60
                    monthly_maintenance_hours[yyyymm] = maintenance_hours
                    
                    # Paradas programadas excluyendo mantención
                    scheduled_mask = scheduled_base_mask & ~maintenance_mask
                    scheduled_data = df.loc[scheduled_mask, duration_col]
                    scheduled_minutes = scheduled_data.sum()
                    scheduled_hours = scheduled_minutes / 60
                    monthly_scheduled_hours[yyyymm] = scheduled_hours

                except Exception as e:
                    print(f"Error procesando {filename}: {e}")
            else:                
                excel_file = pd.ExcelFile(filename)
                if 'detalle' not in excel_file.sheet_names and len(excel_file.sheet_names) > 0:
                    first_sheet = excel_file.sheet_names[0]
                    df = pd.read_excel(filename, sheet_name=first_sheet)
                    duration_col = find_duration_column(df, verbose=False)
                    
                    if duration_col:
                        df[duration_col] = pd.to_numeric(df[duration_col], errors='coerce')
                        total_minutes = df[duration_col].sum()
                        total_hours = total_minutes / 60
                        monthly_hours[yyyymm] = total_hours
                
        except Exception as e:
            print(f"Error leyendo {filename}: {e}")

# Ordenar por fecha
if monthly_hours:
    sorted_months = sorted(monthly_hours.keys())
    
    print("\n=== RESUMEN DE DETENCIONES POR MES ===")
    print("Nota: Todos los valores están en horas\n")
    
    total_year_hours = 0
    total_year_production = 0
    total_year_failures = 0
    total_year_scheduled = 0
    total_year_maintenance = 0
    total_year_unplanned = 0
    total_year_micro_stops = 0
    
    # Mostrar resumen mensual
    current_df = None
    current_duration_col = None
    
    for month in sorted_months:
        total_hours = monthly_hours[month]
        production = monthly_production_hours.get(month, 0)
        failures = monthly_failures_hours.get(month, 0)
        scheduled = monthly_scheduled_hours.get(month, 0)
        maintenance = monthly_maintenance_hours.get(month, 0)
        unplanned = monthly_unplanned_hours.get(month, 0)
        micro_stops = monthly_micro_stops_hours.get(month, 0)
        
        # Actualizar totales anuales
        total_year_hours += total_hours
        total_year_production += production
        total_year_failures += failures
        total_year_scheduled += scheduled
        total_year_maintenance += maintenance
        total_year_unplanned += unplanned
        total_year_micro_stops += micro_stops
        
        # Formatear el mes
        month_date = datetime.strptime(month, '%Y%m')
        month_name = month_date.strftime('%B %Y').capitalize()
        
        print(f"\n{month_name}:")
        print(f"Total horas del mes: {total_hours:.2f} ({total_hours*60:.0f} minutos)")
        
        # Calcular categoría "Otras" (diferencia entre total y suma de categorías)
        categorized_hours = production + failures + maintenance + micro_stops + scheduled + unplanned
        others = total_hours - categorized_hours
        
        # Mostrar horas por categoría
        print(f"\nHoras por categoría:")
        print(f"  Producción: {production:.2f} ({production*60:.0f} minutos)")
        print(f"  Fallas y averías: {failures:.2f} ({failures*60:.0f} minutos)")
        print(f"  Mantenciones: {maintenance:.2f} ({maintenance*60:.0f} minutos)")
        print(f"  Micro paradas: {micro_stops:.2f} ({micro_stops*60:.0f} minutos)")
        print(f"  Paradas programadas (sin mantención): {scheduled:.2f} ({scheduled*60:.0f} minutos)")
        print(f"  Paradas no planificadas (sin fallas): {unplanned:.2f} ({unplanned*60:.0f} minutos)")
        print(f"  Otras: {others:.2f} ({others*60:.0f} minutos)")
        
        # Mostrar porcentajes por categoría
        print(f"\nPorcentajes por categoría:")
        print(f"  Producción: {(production/total_hours*100):.1f}%")
        print(f"  Fallas y averías: {(failures/total_hours*100):.1f}%")
        print(f"  Mantenciones: {(maintenance/total_hours*100):.1f}%")
        print(f"  Micro paradas: {(micro_stops/total_hours*100):.1f}%")
        print(f"  Paradas programadas (sin mantención): {(scheduled/total_hours*100):.1f}%")
        print(f"  Paradas no planificadas (sin fallas): {(unplanned/total_hours*100):.1f}%")
        print(f"  Otras: {(others/total_hours*100):.1f}%")
        
        # Cargar el archivo correspondiente al mes para mostrar ejemplos
        for filename in os.listdir('.'):
            if filename.endswith('.xlsx'):
                match = re.search(pattern, filename)
                if match and match.group(1) == month:
                    try:
                        current_df = pd.read_excel(filename, sheet_name='detalle')
                        
                        # Si las columnas son "Unnamed", intentar con diferentes filas como header
                        if all(col.startswith('Unnamed:') for col in current_df.columns):
                            for header_row in range(0, min(10, len(current_df))):
                                try:
                                    df_test = pd.read_excel(filename, sheet_name='detalle', header=header_row)
                                    duration_col_test = find_duration_column(df_test, verbose=False)
                                    if duration_col_test:
                                        current_df = df_test
                                        current_duration_col = duration_col_test
                                        break
                                except:
                                    continue
                        else:
                            current_duration_col = find_duration_column(current_df, verbose=False)
                        
                        # Mostrar primeros 5 valores de cada categoría para verificar
                        if current_duration_col:
                            print(f"\nPrimeros 5 valores de cada categoría (en minutos):")
                            print(f"  Total: {current_df[current_duration_col].head().tolist()}")
                            
                            mask_prod = current_df['CODIGO DETENCION'].str.lower().isin(['produccion', 'producción'])
                            print(f"  Producción: {current_df.loc[mask_prod, current_duration_col].head().tolist()}")
                            
                            mask_fail = current_df['CODIGO DETENCION'].str.lower().str.contains('fallas y averias|fallas y averías', na=False)
                            print(f"  Fallas y Averías: {current_df.loc[mask_fail, current_duration_col].head().tolist()}")
                            
                            mask_maint = current_df['CODIGO DETENCION'].str.lower().str.contains('mantención|mantencion', na=False)
                            print(f"  Mantenciones Programadas: {current_df.loc[mask_maint, current_duration_col].head().tolist()}")
                            
                            mask_micro = current_df['CODIGO DETENCION'].str.lower().str.contains('micro parada', na=False)
                            print(f"  Micro Paradas: {current_df.loc[mask_micro, current_duration_col].head().tolist()}")
                            
                            mask_sched = current_df['CODIGO DETENCION'].str.lower().str.contains('1. paradas programadas', na=False) & ~current_df['CODIGO DETENCION'].str.lower().str.contains('mantención|mantencion', na=False)
                            print(f"  Paradas Programadas: {current_df.loc[mask_sched, current_duration_col].head().tolist()}")
                            
                            unplanned_base_mask = current_df['CODIGO DETENCION'].str.lower().str.contains('2. paradas no planificadas', na=False)
                            unplanned_mask = unplanned_base_mask & ~mask_fail
                            print(f"  Paradas no planificadas (sin fallas): {current_df.loc[unplanned_mask, current_duration_col].head().tolist()}")
                        
                        break
                    except Exception as e:
                        print(f"Error cargando ejemplos para {month}: {e}")
                        break
        
        print("-"*50)
    
    # Mostrar resumen anual
    # Extraer el año de los datos procesados
    year = sorted_months[0][:4] if sorted_months else "2024"
    print(f"\n=== RESUMEN ANUAL {year} ===")
    print(f"Total horas año: {total_year_hours:.2f} ({total_year_hours*60:.0f} minutos)")
    
    # Calcular categoría "Otras" para el año
    total_year_categorized = total_year_production + total_year_failures + total_year_maintenance + total_year_micro_stops + total_year_scheduled + total_year_unplanned
    total_year_others = total_year_hours - total_year_categorized
    
    # Mostrar horas por categoría
    print(f"\nHoras por categoría:")
    print(f"  Producción: {total_year_production:.2f} ({total_year_production*60:.0f} minutos)")
    print(f"  Fallas y averías: {total_year_failures:.2f} ({total_year_failures*60:.0f} minutos)")
    print(f"  Mantenciones: {total_year_maintenance:.2f} ({total_year_maintenance*60:.0f} minutos)")
    print(f"  Micro paradas: {total_year_micro_stops:.2f} ({total_year_micro_stops*60:.0f} minutos)")
    print(f"  Paradas programadas (sin mantención): {total_year_scheduled:.2f} ({total_year_scheduled*60:.0f} minutos)")
    print(f"  Paradas no planificadas (sin fallas): {total_year_unplanned:.2f} ({total_year_unplanned*60:.0f} minutos)")
    print(f"  Otras: {total_year_others:.2f} ({total_year_others*60:.0f} minutos)")
    
    # Mostrar porcentajes por categoría
    print(f"\nPorcentajes por categoría:")
    print(f"  Producción: {(total_year_production/total_year_hours*100):.1f}%")
    print(f"  Fallas y averías: {(total_year_failures/total_year_hours*100):.1f}%")
    print(f"  Mantenciones: {(total_year_maintenance/total_year_hours*100):.1f}%")
    print(f"  Micro paradas: {(total_year_micro_stops/total_year_hours*100):.1f}%")
    print(f"  Paradas programadas (sin mantención): {(total_year_scheduled/total_year_hours*100):.1f}%")
    print(f"  Paradas no planificadas (sin fallas): {(total_year_unplanned/total_year_hours*100):.1f}%")
    print(f"  Otras: {(total_year_others/total_year_hours*100):.1f}%")