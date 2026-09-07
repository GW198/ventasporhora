import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

# Configuración de la página
st.set_page_config(
    page_title="Reporte de Portabilidades por Hora",
    page_icon="📱",
    layout="wide"
)

# Título principal
st.title("📱 Sistema de Reporte de Portabilidades por Hora")
st.markdown("---")

# Función para procesar los datos
def procesar_portabilidades(df):
    """
    Procesa el DataFrame con las portabilidades
    """
    # Crear columna combinada de fecha y hora
    df['FECHA_HORA'] = pd.to_datetime(df['FECHA_REGISTRO'].astype(str) + ' ' + df['HORA_REGISTRO'].astype(str))
    
    # Extraer la hora y minuto
    df['HORA'] = df['FECHA_HORA'].dt.hour
    df['MINUTO'] = df['FECHA_HORA'].dt.minute
    
    # Función para asignar el rango horario
    def asignar_rango_horario(hora, minuto):
        # Si es exactamente 9:59 o menos, va a las 9
        if hora == 9 and minuto <= 59:
            return '09:00-09:59'
        # Para las demás horas, respeta el rango
        elif hora >= 10 and hora <= 22:
            return f'{hora:02d}:00-{hora:02d}:59'
        else:
            return 'Otros'
    
    # Aplicar la función para asignar rangos
    df['RANGO_HORARIO'] = df.apply(lambda row: asignar_rango_horario(row['HORA'], row['MINUTO']), axis=1)
    
    # Filtrar solo los rangos de interés (9:00 a 22:59)
    df_filtrado = df[df['RANGO_HORARIO'] != 'Otros'].copy()
    
    # Agrupar por rango horario
    reporte = df_filtrado.groupby('RANGO_HORARIO').agg({
        'DN_A_PORTAR': 'count',
        'TIPO_PORTABILIDAD': lambda x: list(x)
    }).reset_index()
    
    reporte.columns = ['RANGO_HORARIO', 'TOTAL_REGISTROS', 'TIPOS_PORTABILIDAD']
    
    # Ordenar los rangos horarios
    orden_rangos = [f'{i:02d}:00-{i:02d}:59' for i in range(9, 23)]
    reporte['RANGO_HORARIO'] = pd.Categorical(reporte['RANGO_HORARIO'], 
                                               categories=orden_rangos, 
                                               ordered=True)
    reporte = reporte.sort_values('RANGO_HORARIO')
    
    # Detalle por tipo de portabilidad
    detalle_por_tipo = df_filtrado.groupby(['RANGO_HORARIO', 'TIPO_PORTABILIDAD']).size().unstack(fill_value=0)
    
    return df_filtrado, reporte, detalle_por_tipo

# Función para crear gráficos
def crear_graficos(reporte, detalle_por_tipo):
    # Gráfico 1: Barras - Total por hora
    fig1 = px.bar(
        reporte,
        x='RANGO_HORARIO',
        y='TOTAL_REGISTROS',
        title='📊 Total de Portabilidades por Hora',
        labels={'RANGO_HORARIO': 'Rango Horario', 'TOTAL_REGISTROS': 'Cantidad'},
        color='TOTAL_REGISTROS',
        color_continuous_scale='Blues'
    )
    fig1.update_layout(height=400)
    
    # Gráfico 2: Líneas - Tendencia por hora
    fig2 = px.line(
        reporte,
        x='RANGO_HORARIO',
        y='TOTAL_REGISTROS',
        title='📈 Tendencia de Portabilidades por Hora',
        labels={'RANGO_HORARIO': 'Rango Horario', 'TOTAL_REGISTROS': 'Cantidad'},
        markers=True
    )
    fig2.update_layout(height=400)
    
    # Gráfico 3: Área apilada - Tipos de portabilidad
    if not detalle_por_tipo.empty:
        fig3 = px.area(
            detalle_por_tipo,
            title='📊 Distribución por Tipo de Portabilidad',
            labels={'value': 'Cantidad', 'RANGO_HORARIO': 'Rango Horario'},
            color_discrete_sequence=px.colors.qualitative.Set2
        )
        fig3.update_layout(height=400)
    else:
        fig3 = None
    
    # Gráfico 4: Pastel - Distribución total
    fig4 = px.pie(
        reporte,
        values='TOTAL_REGISTROS',
        names='RANGO_HORARIO',
        title='🎯 Distribución Porcentual por Hora',
        hole=0.3
    )
    fig4.update_layout(height=400)
    
    return fig1, fig2, fig3, fig4

# Sidebar - Configuración
with st.sidebar:
    st.header("⚙️ Configuración")
    st.markdown("---")
    
    # Subir archivo
    archivo_subido = st.file_uploader(
        "📂 Carga tu archivo Excel",
        type=['xlsx', 'xls'],
        help="Archivo con columnas: USUARIO, NOMBRE_USUARIO, FECHA_REGISTRO, HORA_REGISTRO, DN_A_PORTAR, TIPO_PORTABILIDAD"
    )
    
    st.markdown("---")
    st.info("💡 **Instrucciones:**\n\n"
            "1. Carga tu archivo Excel\n"
            "2. El sistema procesará automáticamente los datos\n"
            "3. Los registros de 9:59 a 0:00 van a la hora 9\n"
            "4. Se generarán reportes y gráficos interactivos")
    
    st.markdown("---")
    st.caption("📱 Desarrollado con Streamlit")

# Área principal
if archivo_subido is not None:
    try:
        # Leer el archivo
        df = pd.read_excel(archivo_subido)
        
        # Verificar columnas
        columnas_requeridas = ['USUARIO', 'NOMBRE_USUARIO', 'FECHA_REGISTRO', 
                              'HORA_REGISTRO', 'DN_A_PORTAR', 'TIPO_PORTABILIDAD']
        
        columnas_faltantes = [col for col in columnas_requeridas if col not in df.columns]
        
        if columnas_faltantes:
            st.error(f"❌ Faltan las siguientes columnas: {', '.join(columnas_faltantes)}")
            st.stop()
        
        # Procesar datos
        with st.spinner('🔄 Procesando datos...'):
            df_filtrado, reporte, detalle_por_tipo = procesar_portabilidades(df)
        
        # Mostrar estadísticas rápidas
        st.success("✅ ¡Datos procesados correctamente!")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📊 Total Registros", len(df))
        with col2:
            st.metric("✅ Procesados", len(df_filtrado))
        with col3:
            st.metric("⏰ Horas Activas", len(reporte))
        with col4:
            max_hora = reporte.loc[reporte['TOTAL_REGISTROS'].idxmax()] if not reporte.empty else None
            if max_hora is not None:
                st.metric("🔥 Hora Pico", max_hora['RANGO_HORARIO'], 
                         f"{max_hora['TOTAL_REGISTROS']} registros")
        
        st.markdown("---")
        
        # Tabs para organización
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Resumen General", "📈 Gráficos", "📋 Detalle", "📥 Exportar"])
        
        with tab1:
            st.subheader("📊 Resumen por Hora")
            
            # Dataframe con resumen
            df_mostrar = reporte.copy()
            df_mostrar['TIPOS_PORTABILIDAD'] = df_mostrar['TIPOS_PORTABILIDAD'].apply(lambda x: ', '.join(set(x)) if isinstance(x, list) else str(x))
            st.dataframe(
                df_mostrar,
                use_container_width=True,
                column_config={
                    "RANGO_HORARIO": "Rango Horario",
                    "TOTAL_REGISTROS": st.column_config.NumberColumn("Total Registros", format="%d"),
                    "TIPOS_PORTABILIDAD": "Tipos de Portabilidad"
                }
            )
        
        with tab2:
            st.subheader("📈 Visualización de Datos")
            
            # Crear gráficos
            fig1, fig2, fig3, fig4 = crear_graficos(reporte, detalle_por_tipo)
            
            # Mostrar gráficos en 2 columnas
            col1, col2 = st.columns(2)
            with col1:
                st.plotly_chart(fig1, use_container_width=True)
            with col2:
                st.plotly_chart(fig2, use_container_width=True)
            
            col3, col4 = st.columns(2)
            with col3:
                if fig3:
                    st.plotly_chart(fig3, use_container_width=True)
                else:
                    st.info("No hay suficientes datos para mostrar la distribución por tipo")
            with col4:
                st.plotly_chart(fig4, use_container_width=True)
        
        with tab3:
            st.subheader("📋 Detalle de Registros")
            
            # Mostrar detalle por tipo
            st.write("**Distribución por Tipo de Portabilidad:**")
            st.dataframe(
                detalle_por_tipo,
                use_container_width=True
            )
            
            st.markdown("---")
            
            # Mostrar todos los registros procesados
            st.write("**Todos los Registros Procesados:**")
            df_completo = df_filtrado[['USUARIO', 'NOMBRE_USUARIO', 'FECHA_REGISTRO', 
                                      'HORA_REGISTRO', 'DN_A_PORTAR', 'TIPO_PORTABILIDAD', 
                                      'RANGO_HORARIO']]
            st.dataframe(
                df_completo,
                use_container_width=True,
                height=400
            )
        
        with tab4:
            st.subheader("📥 Exportar Reportes")
            
            col1, col2 = st.columns(2)
            
            with col1:
                # Exportar reporte
                output = pd.ExcelWriter('reporte_temp.xlsx', engine='openpyxl')
                reporte.to_excel(output, sheet_name='Resumen_Hora', index=False)
                detalle_por_tipo.to_excel(output, sheet_name='Detalle_Tipo')
                df_filtrado[['USUARIO', 'NOMBRE_USUARIO', 'FECHA_REGISTRO', 
                            'HORA_REGISTRO', 'DN_A_PORTAR', 'TIPO_PORTABILIDAD', 
                            'RANGO_HORARIO']].to_excel(output, sheet_name='Registros_Completos', index=False)
                output.close()
                
                with open('reporte_temp.xlsx', 'rb') as f:
                    st.download_button(
                        label="📥 Descargar Reporte Excel",
                        data=f,
                        file_name=f'reporte_portabilidades_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
                # Eliminar archivo temporal
                os.remove('reporte_temp.xlsx')
            
            with col2:
                # Exportar CSV
                csv = reporte.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Descargar CSV",
                    data=csv,
                    file_name=f'reporte_portabilidades_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
                    mime='text/csv'
                )
            
            st.markdown("---")
            
            # Vista previa de datos a exportar
            st.write("**Vista previa del reporte a exportar:**")
            st.dataframe(reporte, use_container_width=True)
    
    except Exception as e:
        st.error(f"❌ Error al procesar el archivo: {str(e)}")
        st.info("Verifica que el archivo tenga el formato correcto y las columnas necesarias.")

else:
    # Mensaje cuando no hay archivo
    st.markdown("### 👈 Carga un archivo Excel para comenzar")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.info("📋 **Requisitos del archivo:**\n\n"
                "• Formato: .xlsx o .xls\n"
                "• Columnas requeridas:\n"
                "  - USUARIO\n"
                "  - NOMBRE_USUARIO\n"
                "  - FECHA_REGISTRO\n"
                "  - HORA_REGISTRO\n"
                "  - DN_A_PORTAR\n"
                "  - TIPO_PORTABILIDAD")
    
    with col2:
        st.warning("⏰ **Reglas de agrupación:**\n\n"
                   "• 9:00 a 9:59 → Hora 9\n"
                   "• 10:00 a 10:59 → Hora 10\n"
                   "• ... y así sucesivamente\n"
                   "• 9:59 a 0:00 → Hora 9")
    
    with col3:
        st.success("📊 **Reportes generados:**\n\n"
                   "• Resumen por hora\n"
                   "• Gráficos interactivos\n"
                   "• Detalle por tipo\n"
                   "• Exportación Excel/CSV")

st.markdown("---")
st.caption("© 2024 - Sistema de Reporte de Portabilidades | Desarrollado con ❤️ usando Streamlit")
