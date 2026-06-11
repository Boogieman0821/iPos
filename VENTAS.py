import pandas as pd

# ==============================
# CONFIGURACIÓN
# ==============================
file_path = r"C:\Users\jimgl\Downloads\VENTAS.xlsx"
output_path = r"C:\Users\jimgl\Downloads\VENTAS_Procesadas.xlsx"

# ==============================
# 1. Cargar las hojas
# ==============================
xls = pd.ExcelFile(file_path)
ventas_df = pd.read_excel(xls, sheet_name='Ventas')
orden_df = pd.read_excel(xls, sheet_name='ORDEN')

# ==============================
# 2. Obtener drivers únicos
# ==============================
drivers = orden_df['driver'].unique()

# ==============================
# 3. Crear ExcelWriter para exportar resultados
# ==============================
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    
    for driver in drivers:
        # ------------------------------
        # Paso 1. Filtrar driver actual
        # ------------------------------
        orden_driver_df = orden_df[orden_df['driver'] == driver]
        
        # ------------------------------
        # Paso 2. Crear tabla 'paradas' temporal
        # ------------------------------
        paradas_df = orden_driver_df[['stop_number', 'external_id']].copy()
        paradas_df.columns = ['#', 'cliente']  # Renombrar columnas como en la tabla paradas
        
        # ------------------------------
        # Paso 3. Simular BUSCARV: mapear external_id a stop_number
        # ------------------------------
        lookup_dict = dict(zip(paradas_df['cliente'], paradas_df['#']))
        
        # Crear columna E (#) con el resultado del lookup
        ventas_df['#'] = ventas_df['Num Cliente'].map(lookup_dict)
        
        # ------------------------------
        # Paso 4. Filtrar eliminando NAs y ordenar por columna E (#)
        # ------------------------------
        ventas_filtradas = ventas_df.dropna(subset=['#'])
        ventas_filtradas = ventas_filtradas.sort_values(by='#')
        
        # ------------------------------
        # Paso 5. Guardar la tabla filtrada en nueva hoja con nombre del driver
        # ------------------------------
        ventas_filtradas.to_excel(writer, sheet_name=str(driver)[:31], index=False)
        
        print(f"✅ Tabla generada para driver: {driver}")
        
        # ------------------------------
        # Paso 6. Limpiar columna E (#) antes de procesar el siguiente driver
        # ------------------------------
        ventas_df['#'] = None

# ==============================
# 4. Confirmación final
# ==============================
print(f"🎉 Archivo generado exitosamente en: {output_path}")
