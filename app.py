import streamlit as st
import pandas as pd
import io
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

# Configuración de la página
st.set_page_config(
    page_title="Reporte de Portabilidades por Hora",
    page_icon="📱",
    layout="wide"
)

# Estilo personalizado
st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
    }
    </style>
""", unsafe_allow_html=True)

# Título
st.markdown('<div class="main-header">📱 Sistema de Reporte de Portabilidades</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Carga tu archivo Excel y obtén reportes automáticos por hora</div>', unsafe_allow_html=True)

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

# Función para crear gráficos
def crear_graficos(reporte, detalle_tipo):
    try:
        # Gráfico 1: Barras
        fig1 = px.bar(
            reporte,
            x='RANGO_HORARIO',
            y='TOTAL_REGISTROS',
            title='Total de Portabilidades por Hora',
            labels={'RANGO_HORARIO': 'Rango Horario', 'TOTAL_REGISTROS': 'Cantidad'},
            color='TOTAL_REGISTROS',
            color_continuous_scale='Blues',
            text='TOTAL_REGISTROS'
        )
        fig1.update_traces(texttemplate='%{text}', textposition='outside')
        fig1.update_layout(height=400, showlegend=False)
        
        # Gráfico 2: Líneas
        fig2 = px.line(
            reporte,
            x='RANGO_HORARIO',
            y='TOTAL_REGISTROS',
            title='Tendencia de Portabilidades',
            labels={'RANGO_HORARIO': 'Rango Horario', 'TOTAL_REGISTROS': 'Cantidad'},
            markers=True
        )
        fig2.update_layout(height=400)
        
        # Gráfico 3: Área
        if not detalle_tipo.empty:
            fig3 = px.area(
                detalle_tipo,
                title='Distribución por Tipo de Portabilidad',
                labels={'value': 'Cantidad', 'RANGO_HORARIO': 'Rango Horario'},
                color_discrete_sequence=px.colors.qualitative.Set2
            )
            fig3.update_layout(height=400)
        else:
            fig3 = None
        
        # Gráfico 4: Pastel
        fig4 = px.pie(
            reporte,
            values='TOTAL_REGISTROS',
            names='RANGO_HORARIO',
            title='Distribución Porcentual',
            hole=0.3
        )
        fig4.update_layout(height=400)
        
        return fig1, fig2, fig3, fig4
        
    except Exception as e:
        st.error(f"Error al crear gráficos: {str(e)}")
        return None, None, None, None

# Interfaz principal
col1, col2, col3 = st.columns([1, 2, 1])

with col2:
    st.markdown("### 📂 Carga tu archivo")
    
    archivo = st.file_uploader(
        "Selecciona un archivo Excel",
        type=['xlsx', 'xls'],
        help="El archivo debe contener las columnas requeridas",
        label_visibility="collapsed"
    )

# Si hay archivo cargado
if archivo is not None:
    try:
        # Leer archivo
        with st.spinner('📖 Leyendo archivo...'):
            df = pd.read_excel(archivo)
        
        # Mostrar info del archivo
        st.success(f"✅ Archivo '{archivo.name}' cargado correctamente")
        
        col_info1, col_info2, col_info3 = st.columns(3)
        with col_info1:
            st.metric("📄 Registros", len(df))
        with col_info2:
            st.metric("📋 Columnas", len(df.columns))
        with col_info3:
            st.metric("📊 Tamaño", f"{archivo.size / 1024:.1f} KB")
        
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
        
        # Métricas
        st.markdown("---")
        st.subheader("📊 Resumen Ejecutivo")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Registros", f"{len(df):,}")
        with col2:
            st.metric("Registros Procesados", f"{len(df_filtrado):,}")
        with col3:
            horas_activas = len(reporte[reporte['TOTAL_REGISTROS'] > 0])
            st.metric("Horas con Actividad", horas_activas)
        with col4:
            if not reporte.empty:
                max_hora = reporte.loc[reporte['TOTAL_REGISTROS'].idxmax()]
                st.metric("🚀 Hora Pico", max_hora['RANGO_HORARIO'], 
                         f"{max_hora['TOTAL_REGISTROS']} registros")
        
        # Tabs
        tab1, tab2, tab3 = st.tabs(["📊 Resumen por Hora", "📈 Visualizaciones", "📋 Detalle Completo"])
        
        with tab1:
            st.subheader("Reporte por Rango Horario")
            
            # Mostrar tabla
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
            st.markdown("---")
            col1, col2 = st.columns(2)
            with col1:
                st.write("**Estadísticas:**")
                stats = reporte['TOTAL_REGISTROS'].describe()
                stats_df = pd.DataFrame({
                    'Estadística': ['Total', 'Promedio', 'Mínimo', 'Máximo', 'Desviación'],
                    'Valor': [
                        f"{stats['count']:.0f}",
                        f"{stats['mean']:.1f}",
                        f"{stats['min']:.0f}",
                        f"{stats['max']:.0f}",
                        f"{stats['std']:.1f}"
                    ]
                })
                st.dataframe(stats_df, hide_index=True)
            with col2:
                if not detalle_tipo.empty:
                    st.write("**Distribución por Tipo:**")
                    total_por_tipo = detalle_tipo.sum().sort_values(ascending=False)
                    st.dataframe(
                        pd.DataFrame({
                            'Tipo': total_por_tipo.index,
                            'Total': total_por_tipo.values
                        }),
                        hide_index=True
                    )
        
        with tab2:
            st.subheader("Visualizaciones Interactivas")
            
            # Crear gráficos
            fig1, fig2, fig3, fig4 = crear_graficos(reporte, detalle_tipo)
            
            # Mostrar gráficos
            col1, col2 = st.columns(2)
            with col1:
                if fig1:
                    st.plotly_chart(fig1, use_container_width=True)
            with col2:
                if fig2:
                    st.plotly_chart(fig2, use_container_width=True)
            
            col3, col4 = st.columns(2)
            with col3:
                if fig3:
                    st.plotly_chart(fig3, use_container_width=True)
                else:
                    st.info("No hay datos suficientes para el gráfico de área")
            with col4:
                if fig4:
                    st.plotly_chart(fig4, use_container_width=True)
        
        with tab3:
            st.subheader("Detalle de Todos los Registros")
            
            # Mostrar todos los registros
            df_completo = df_filtrado[[
                'USUARIO', 'NOMBRE_USUARIO', 'FECHA_REGISTRO', 
                'HORA_REGISTRO', 'DN_A_PORTAR', 'TIPO_PORTABILIDAD', 
                'RANGO_HORARIO'
            ]]
            
            st.dataframe(
                df_completo,
                use_container_width=True,
                height=500
            )
            
            # Opciones de descarga
            st.markdown("---")
            st.subheader("📥 Descargar Reportes")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                # Descargar Excel
                try:
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        reporte.to_excel(writer, sheet_name='Resumen_Hora', index=False)
                        if not detalle_tipo.empty:
                            detalle_tipo.to_excel(writer, sheet_name='Detalle_Tipo')
                        df_completo.to_excel(writer, sheet_name='Registros', index=False)
                    
                    output.seek(0)
                    st.download_button(
                        label="📥 Excel",
                        data=output,
                        file_name=f'reporte_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Error: {str(e)}")
            
            with col2:
                # Descargar CSV
                try:
                    csv = reporte.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 CSV",
                        data=csv,
                        file_name=f'reporte_{datetime.now().strftime("%Y%m%d_%H%M")}.csv',
                        mime='text/csv',
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Error: {str(e)}")
            
            with col3:
                # Descargar JSON
                try:
                    json_data = reporte.to_json(orient='records', date_format='iso')
                    st.download_button(
                        label="📥 JSON",
                        data=json_data,
                        file_name=f'reporte_{datetime.now().strftime("%Y%m%d_%H%M")}.json',
                        mime='application/json',
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Error: {str(e)}")
        
    except Exception as e:
        st.error(f"❌ Error inesperado: {str(e)}")
        st.stop()

else:
    # Mensaje inicial
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.info("""
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
        st.warning("""
        ### ⏰ Reglas
        - 9:00 a 9:59 → Hora 9
        - 10:00 a 10:59 → Hora 10
        - 11:00 a 11:59 → Hora 11
        - ... y así sucesivamente
        - 9:59 a 0:00 → Hora 9
        """)
    
    with col3:
        st.success("""
        ### 📊 Reportes
        - Resumen por hora
        - Gráficos interactivos
        - Detalle por tipo
        - Exportación Excel/CSV/JSON
        - Estadísticas automáticas
        """)

# Footer
st.markdown("---")
st.caption("💡 Desarrollado con Streamlit | Carga tu archivo y obtén reportes automáticos")
