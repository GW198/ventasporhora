import streamlit as st
import pandas as pd
import io
from datetime import datetime

# Configuración de la página
st.set_page_config(
    page_title="Reporte de Portabilidades por Hora",
    page_icon="📱",
    layout="wide"
)

# Título
st.title("📱 Sistema de Reporte de Portabilidades por Hora")
st.markdown("---")

# Función para procesar los datos
def procesar_portabilidades(df):
    """Procesa el DataFrame con las portabilidades"""
    try:
        # Crear copia
        df_procesado = df.copy()
        
        # Convertir fecha
        df_procesado['FECHA_REGISTRO'] = pd.to_datetime(df_procesado['FECHA_REGISTRO']).dt.date
        
        # Convertir hora - soporta múltiples formatos
        try:
            df_procesado['HORA_REGISTRO'] = pd.to_datetime(df_procesado['HORA_REGISTRO'].astype(str), format='%H:%M:%S').dt.time
        except:
            try:
                df_procesado['HORA_REGISTRO'] = pd.to_datetime(df_procesado['HORA_REGISTRO'].astype(str), format='%H:%M').dt.time
            except:
                df_procesado['HORA_REGISTRO'] = pd.to_datetime(df_procesado['HORA_REGISTRO'].astype(str)).dt.time
        
        # Extraer hora y minuto
        df_procesado['HORA'] = df_procesado['HORA_REGISTRO'].apply(lambda x: x.hour)
        df_procesado['MINUTO'] = df_procesado['HORA_REGISTRO'].apply(lambda x: x.minute)
        
        # Asignar rango horario
        def asignar_rango(hora, minuto):
            if hora == 9 and minuto <= 59:
                return '09:00-09:59'
            elif 10 <= hora <= 22:
                return f'{hora:02d}:00-{hora:02d}:59'
            else:
                return 'Otros'
        
        df_procesado['RANGO_HORARIO'] = df_procesado.apply(
            lambda row: asignar_rango(row['HORA'], row['MINUTO']), axis=1
        )
        
        # Filtrar rangos válidos
        df_filtrado = df_procesado[df_procesado['RANGO_HORARIO'] != 'Otros'].copy()
        
        if df_filtrado.empty:
            return None, None, None, "No se encontraron registros en el rango de 9:00 a 22:59"
        
        # Crear reporte por hora
        reporte = df_filtrado.groupby('RANGO_HORARIO').agg({
            'DN_A_PORTAR': 'count',
            'TIPO_PORTABILIDAD': lambda x: list(x)
        }).reset_index()
        
        reporte.columns = ['RANGO_HORARIO', 'TOTAL_REGISTROS', 'TIPOS_PORTABILIDAD']
        
        # Ordenar
        orden = [f'{i:02d}:00-{i:02d}:59' for i in range(9, 23)]
        reporte['RANGO_HORARIO'] = pd.Categorical(reporte['RANGO_HORARIO'], categories=orden, ordered=True)
        reporte = reporte.sort_values('RANGO_HORARIO').reset_index(drop=True)
        
        # Detalle por tipo
        detalle_tipo = df_filtrado.groupby(['RANGO_HORARIO', 'TIPO_PORTABILIDAD']).size().unstack(fill_value=0)
        
        return df_filtrado, reporte, detalle_tipo, None
        
    except Exception as e:
        return None, None, None, f"Error al procesar: {str(e)}"

# Interfaz de carga de archivo
st.markdown("### 📂 Carga tu archivo Excel")
st.info("Asegúrate de que tu archivo tenga las columnas: USUARIO, NOMBRE_USUARIO, FECHA_REGISTRO, HORA_REGISTRO, DN_A_PORTAR, TIPO_PORTABILIDAD")

archivo = st.file_uploader(
    "Selecciona un archivo Excel",
    type=['xlsx', 'xls'],
    help="El archivo debe contener las columnas requeridas",
    label_visibility="collapsed"
)

if archivo is not None:
    try:
        # Leer archivo
        with st.spinner('📖 Leyendo archivo...'):
            df = pd.read_excel(archivo)
        
        st.success(f"✅ Archivo '{archivo.name}' cargado correctamente")
        
        # Verificar columnas
        columnas_requeridas = ['USUARIO', 'NOMBRE_USUARIO', 'FECHA_REGISTRO', 
                              'HORA_REGISTRO', 'DN_A_PORTAR', 'TIPO_PORTABILIDAD']
        
        columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]
        
        if columnas_faltantes:
            st.error(f"❌ Faltan columnas: {', '.join(columnas_faltantes)}")
            st.info("Las columnas requeridas son: " + ", ".join(columnas_requeridas))
            st.stop()
        
        # Procesar datos
        with st.spinner('🔄 Procesando datos...'):
            df_filtrado, reporte, detalle_tipo, error = procesar_portabilidades(df)
        
        if error:
            st.error(f"❌ {error}")
            st.stop()
        
        if df_filtrado is None or reporte.empty:
            st.warning("⚠️ No se encontraron registros para procesar")
            st.stop()
        
        # Mostrar estadísticas
        st.markdown("---")
        st.subheader("📊 Resumen Ejecutivo")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📄 Total Registros", f"{len(df):,}")
        with col2:
            st.metric("✅ Procesados", f"{len(df_filtrado):,}")
        with col3:
            horas_activas = len(reporte[reporte['TOTAL_REGISTROS'] > 0])
            st.metric("⏰ Horas Activas", horas_activas)
        with col4:
            if not reporte.empty:
                max_hora = reporte.loc[reporte['TOTAL_REGISTROS'].idxmax()]
                st.metric("🚀 Hora Pico", max_hora['RANGO_HORARIO'], 
                         f"{max_hora['TOTAL_REGISTROS']} registros")
        
        # Mostrar reporte
        st.markdown("---")
        st.subheader("📊 Reporte por Rango Horario")
        
        # Tabla de reporte
        df_mostrar = reporte.copy()
        df_mostrar['TIPOS_PORTABILIDAD'] = df_mostrar['TIPOS_PORTABILIDAD'].apply(
            lambda x: ', '.join(set(x)) if isinstance(x, list) else str(x)
        )
        
        st.dataframe(
            df_mostrar,
            use_container_width=True,
            column_config={
                "RANGO_HORARIO": "Rango Horario",
                "TOTAL_REGISTROS": st.column_config.NumberColumn("Total", format="%d"),
                "TIPOS_PORTABILIDAD": "Tipos de Portabilidad"
            }
        )
        
        # Estadísticas adicionales
        col1, col2 = st.columns(2)
        with col1:
            st.write("**📈 Estadísticas de la distribución:**")
            stats = reporte['TOTAL_REGISTROS'].describe()
            stats_df = pd.DataFrame({
                'Estadística': ['Total', 'Promedio', 'Mínimo', 'Máximo', 'Mediana', 'Desviación'],
                'Valor': [
                    f"{stats['count']:.0f}",
                    f"{stats['mean']:.1f}",
                    f"{stats['min']:.0f}",
                    f"{stats['max']:.0f}",
                    f"{stats['50%']:.1f}",
                    f"{stats['std']:.1f}"
                ]
            })
            st.dataframe(stats_df, hide_index=True, use_container_width=True)
        
        with col2:
            if not detalle_tipo.empty:
                st.write("**📊 Distribución por Tipo:**")
                total_por_tipo = detalle_tipo.sum().sort_values(ascending=False)
                st.dataframe(
                    pd.DataFrame({
                        'Tipo': total_por_tipo.index,
                        'Total': total_por_tipo.values
                    }),
                    hide_index=True,
                    use_container_width=True
                )
        
        # Detalle completo
        with st.expander("📋 Ver todos los registros procesados"):
            df_completo = df_filtrado[[
                'USUARIO', 'NOMBRE_USUARIO', 'FECHA_REGISTRO', 
                'HORA_REGISTRO', 'DN_A_PORTAR', 'TIPO_PORTABILIDAD', 
                'RANGO_HORARIO'
            ]]
            st.dataframe(df_completo, use_container_width=True, height=400)
        
        # Exportar
        st.markdown("---")
        st.subheader("📥 Exportar Reportes")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            # Exportar Excel
            try:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    reporte.to_excel(writer, sheet_name='Resumen_Hora', index=False)
                    if not detalle_tipo.empty:
                        detalle_tipo.to_excel(writer, sheet_name='Detalle_Tipo')
                    if not df_filtrado.empty:
                        df_filtrado[['USUARIO', 'NOMBRE_USUARIO', 'FECHA_REGISTRO', 
                                    'HORA_REGISTRO', 'DN_A_PORTAR', 'TIPO_PORTABILIDAD', 
                                    'RANGO_HORARIO']].to_excel(writer, sheet_name='Registros', index=False)
                
                output.seek(0)
                st.download_button(
                    label="📥 Descargar Excel",
                    data=output,
                    file_name=f'reporte_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Error al generar Excel: {str(e)}")
        
        with col2:
            # Exportar CSV
            try:
                csv = reporte.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Descargar CSV",
                    data=csv,
                    file_name=f'reporte_{datetime.now().strftime("%Y%m%d_%H%M")}.csv',
                    mime='text/csv',
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Error al generar CSV: {str(e)}")
        
        with col3:
            # Exportar JSON
            try:
                json_data = reporte.to_json(orient='records', date_format='iso')
                st.download_button(
                    label="📥 Descargar JSON",
                    data=json_data,
                    file_name=f'reporte_{datetime.now().strftime("%Y%m%d_%H%M")}.json',
                    mime='application/json',
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Error al generar JSON: {str(e)}")
    
    except Exception as e:
        st.error(f"❌ Error inesperado: {str(e)}")
        st.stop()

else:
    # Mensaje inicial
    st.markdown("---")
    st.info("👆 Carga un archivo Excel para comenzar")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        ### 📋 Requisitos
        - Archivo Excel (.xlsx o .xls)
        - Columnas requeridas:
          - USUARIO
          - NOMBRE_USUARIO
          - FECHA_REGISTRO
          - HORA_REGISTRO
          - DN_A_PORTAR
          - TIPO_PORTABILIDAD
        """)
    
    with col2:
        st.markdown("""
        ### ⏰ Reglas de agrupación
        - 9:00 a 9:59 → Hora 9
        - 10:00 a 10:59 → Hora 10
        - 11:00 a 11:59 → Hora 11
        - ... y así sucesivamente
        - 9:59 a 0:00 → Hora 9
        """)
    
    with col3:
        st.markdown("""
        ### 📊 Reportes generados
        - Resumen por hora
        - Estadísticas automáticas
        - Detalle por tipo
        - Exportación Excel/CSV/JSON
        - Visualización de datos
        """)

# Footer
st.markdown("---")
st.caption("💡 Sistema de Reporte de Portabilidades | Carga tu archivo y obtén análisis automáticos")
