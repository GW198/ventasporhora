import pandas as pd
import os
from datetime import datetime, time

def procesar_portabilidades(archivo_excel):
    """
    Procesa un archivo Excel con portabilidades y genera reporte por horas
    
    Args:
        archivo_excel (str): Ruta del archivo Excel
    """
    
    # Leer el archivo Excel
    try:
        df = pd.read_excel(archivo_excel)
        print(f"✅ Archivo cargado correctamente. {len(df)} registros encontrados.")
    except Exception as e:
        print(f"❌ Error al leer el archivo: {e}")
        return
    
    # Verificar que las columnas existan
    columnas_requeridas = ['USUARIO', 'NOMBRE_USUARIO', 'FECHA_REGISTRO', 
                          'HORA_REGISTRO', 'DN_A_PORTAR', 'TIPO_PORTABILIDAD']
    
    for col in columnas_requeridas:
        if col not in df.columns:
            print(f"❌ Error: La columna '{col}' no existe en el archivo")
            return
    
    # Crear columna combinada de fecha y hora
    df['FECHA_HORA'] = pd.to_datetime(df['FECHA_REGISTRO'].astype(str) + ' ' + df['HORA_REGISTRO'].astype(str))
    
    # Extraer la hora
    df['HORA'] = df['FECHA_HORA'].dt.hour
    df['MINUTO'] = df['FECHA_HORA'].dt.minute
    
    # Función para asignar el rango horario
    def asignar_rango_horario(hora, minuto):
        # Si es exactamente 9:59 o menos, va a las 9
        if hora == 9 and minuto <= 59:
            return '09:00-09:59'
        # Para las demás horas, respeta el rango
        elif hora >= 10 and hora <= 22:  # Limitamos hasta las 22:59
            return f'{hora:02d}:00-{hora:02d}:59'
        else:
            return 'Otros'
    
    # Aplicar la función para asignar rangos
    df['RANGO_HORARIO'] = df.apply(lambda row: asignar_rango_horario(row['HORA'], row['MINUTO']), axis=1)
    
    # Filtrar solo los rangos de interés (9:00 a 22:59)
    df_filtrado = df[df['RANGO_HORARIO'] != 'Otros'].copy()
    
    # Agrupar por rango horario y contar los DN_A_PORTAR (columna D)
    reporte = df_filtrado.groupby('RANGO_HORARIO').agg({
        'DN_A_PORTAR': 'count',  # Contar registros
        'TIPO_PORTABILIDAD': lambda x: list(x)  # Lista de tipos de portabilidad
    }).reset_index()
    
    # Renombrar columnas
    reporte.columns = ['RANGO_HORARIO', 'TOTAL_REGISTROS', 'TIPOS_PORTABILIDAD']
    
    # Ordenar los rangos horarios
    orden_rangos = [f'{i:02d}:00-{i:02d}:59' for i in range(9, 23)]
    reporte['RANGO_HORARIO'] = pd.Categorical(reporte['RANGO_HORARIO'], 
                                               categories=orden_rangos, 
                                               ordered=True)
    reporte = reporte.sort_values('RANGO_HORARIO')
    
    # Crear resumen detallado por tipo de portabilidad
    detalle_por_tipo = df_filtrado.groupby(['RANGO_HORARIO', 'TIPO_PORTABILIDAD']).size().unstack(fill_value=0)
    
    # Mostrar resultados
    print("\n" + "="*60)
    print("📊 REPORTE DE PORTABILIDADES POR HORA")
    print("="*60)
    print("\n📋 Resumen por hora:")
    print(reporte.to_string(index=False))
    
    print("\n📈 Detalle por tipo de portabilidad:")
    print(detalle_por_tipo)
    
    # Generar archivo de salida
    nombre_base = os.path.splitext(archivo_excel)[0]
    archivo_salida = f"{nombre_base}_reporte_por_hora.xlsx"
    
    with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
        # Hoja 1: Resumen
        reporte.to_excel(writer, sheet_name='Resumen_Hora', index=False)
        
        # Hoja 2: Detalle por tipo
        detalle_por_tipo.to_excel(writer, sheet_name='Detalle_Tipo')
        
        # Hoja 3: Todos los registros procesados
        df_filtrado[['USUARIO', 'NOMBRE_USUARIO', 'FECHA_REGISTRO', 
                    'HORA_REGISTRO', 'DN_A_PORTAR', 'TIPO_PORTABILIDAD', 
                    'RANGO_HORARIO']].to_excel(writer, sheet_name='Registros_Completos', index=False)
    
    print(f"\n✅ Reporte guardado en: {archivo_salida}")
    
    # Estadísticas adicionales
    print("\n📊 ESTADÍSTICAS ADICIONALES:")
    print(f"Total de registros procesados: {len(df_filtrado)}")
    print(f"Total de registros excluidos (fuera de rango): {len(df) - len(df_filtrado)}")
    
    # Mostrar hora con más registros
    if not reporte.empty:
        max_hora = reporte.loc[reporte['TOTAL_REGISTROS'].idxmax()]
        print(f"Hora con más portabilidades: {max_hora['RANGO_HORARIO']} ({max_hora['TOTAL_REGISTROS']} registros)")
    
    return reporte, detalle_por_tipo

# Función principal para ejecutar el programa
def main():
    print("="*60)
    print("📱 SISTEMA DE REPORTE DE PORTABILIDADES POR HORA")
    print("="*60)
    
    # Solicitar la ruta del archivo
    while True:
        archivo = input("\n📂 Ingresa la ruta del archivo Excel: ").strip()
        
        if os.path.exists(archivo) and archivo.endswith(('.xlsx', '.xls')):
            break
        else:
            print("❌ El archivo no existe o no es un archivo Excel válido. Intenta de nuevo.")
    
    # Procesar el archivo
    procesar_portabilidades(archivo)

# Ejemplo de uso directo (comentar si usas main())
if __name__ == "__main__":
    # Opción 1: Usar la función main() para interactuar
    main()
    
    # Opción 2: Llamar directamente con la ruta del archivo
    # procesar_portabilidades("ruta_del_archivo.xlsx")
