""")

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
# Leer el archivo con manejo de errores
with st.spinner('🔄 Leyendo archivo...'):
    try:
        if archivo_subido.name.endswith('.xlsx'):
            df = pd.read_excel(archivo_subido, engine='openpyxl')
        else:
            df = pd.read_excel(archivo_subido)
    except Exception as e:
        st.error(f"❌ Error al leer el archivo: {str(e)}")
        st.info("Asegúrate de que el archivo sea un Excel válido y no esté corrupto.")
        st.stop()

# Mostrar información del archivo
col_info1, col_info2, col_info3 = st.columns(3)
with col_info1:
    st.info(f"📄 Archivo: {archivo_subido.name}")
with col_info2:
    st.info(f"📊 Filas: {len(df)}")
with col_info3:
    st.info(f"📋 Columnas: {len(df.columns)}")

# Procesar datos
with st.spinner('🔄 Procesando datos...'):
    df_filtrado, reporte, detalle_por_tipo = procesar_portabilidades(df)

if df_filtrado is None:
    st.stop()

if df_filtrado.empty:
    st.warning("⚠️ No se encontraron registros para procesar. Verifica que los datos cumplan con el formato esperado.")
    st.stop()

# Mostrar estadísticas rápidas
st.success("✅ ¡Datos procesados correctamente!")

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("📊 Total Registros", f"{len(df):,}")
with col2:
    st.metric("✅ Procesados", f"{len(df_filtrado):,}")
with col3:
    horas_activas = len(reporte[reporte['TOTAL_REGISTROS'] > 0]) if not reporte.empty else 0
    st.metric("⏰ Horas Activas", horas_activas)
with col4:
    if not reporte.empty:
        max_hora = reporte.loc[reporte['TOTAL_REGISTROS'].idxmax()]
        st.metric("🔥 Hora Pico", max_hora['RANGO_HORARIO'], 
                 f"{max_hora['TOTAL_REGISTROS']:,} registros")
    else:
        st.metric("🔥 Hora Pico", "Sin datos")

st.markdown("---")

# Tabs para organización
tab1, tab2, tab3, tab4 = st.tabs(["📊 Resumen General", "📈 Gráficos", "📋 Detalle", "📥 Exportar"])

with tab1:
    st.subheader("📊 Resumen por Hora")
    
    if not reporte.empty:
        # Dataframe con resumen
        df_mostrar = reporte.copy()
        df_mostrar['TIPOS_PORTABILIDAD'] = df_mostrar['TIPOS_PORTABILIDAD'].apply(
            lambda x: ', '.join(set(x)) if isinstance(x, list) and len(x) > 0 else 'Sin datos'
        )
        st.dataframe(
            df_mostrar,
            use_container_width=True,
            column_config={
                "RANGO_HORARIO": "Rango Horario",
                "TOTAL_REGISTROS": st.column_config.NumberColumn("Total Registros", format="%d"),
                "TIPOS_PORTABILIDAD": "Tipos de Portabilidad"
            }
        )
    else:
        st.info("No hay datos para mostrar")

with tab2:
    st.subheader("📈 Visualización de Datos")
    
    if not reporte.empty:
        # Crear gráficos
        fig1, fig2, fig3, fig4 = crear_graficos(reporte, detalle_por_tipo)
        
        # Mostrar gráficos en 2 columnas
        col1, col2 = st.columns(2)
        with col1:
            if fig1:
                st.plotly_chart(fig1, use_container_width=True)
            else:
                st.info("No hay datos para mostrar el gráfico de barras")
        with col2:
            if fig2:
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.info("No hay datos para mostrar el gráfico de líneas")
        
        col3, col4 = st.columns(2)
        with col3:
            if fig3:
                st.plotly_chart(fig3, use_container_width=True)
            else:
                st.info("No hay suficientes datos para mostrar la distribución por tipo")
        with col4:
            if fig4:
                st.plotly_chart(fig4, use_container_width=True)
            else:
                st.info("No hay datos para mostrar el gráfico de pastel")
    else:
        st.info("No hay datos suficientes para generar gráficos")

with tab3:
    st.subheader("📋 Detalle de Registros")
    
    # Mostrar detalle por tipo
    st.write("**Distribución por Tipo de Portabilidad:**")
    if not detalle_por_tipo.empty:
        st.dataframe(
            detalle_por_tipo,
            use_container_width=True
        )
    else:
        st.info("No hay datos para mostrar la distribución por tipo")
    
    st.markdown("---")
    
    # Mostrar todos los registros procesados
    st.write("**Todos los Registros Procesados:**")
    if not df_filtrado.empty:
        df_completo = df_filtrado[['USUARIO', 'NOMBRE_USUARIO', 'FECHA_REGISTRO', 
                                  'HORA_REGISTRO', 'DN_A_PORTAR', 'TIPO_PORTABILIDAD', 
                                  'RANGO_HORARIO']]
        st.dataframe(
            df_completo,
            use_container_width=True,
            height=400
        )
    else:
        st.info("No hay registros para mostrar")

with tab4:
    st.subheader("📥 Exportar Reportes")
    
    if not reporte.empty:
        col1, col2 = st.columns(2)
        
        with col1:
            # Exportar Excel
            try:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    reporte.to_excel(writer, sheet_name='Resumen_Hora', index=False)
                    if not detalle_por_tipo.empty:
                        detalle_por_tipo.to_excel(writer, sheet_name='Detalle_Tipo')
                    if not df_filtrado.empty:
                        df_filtrado[['USUARIO', 'NOMBRE_USUARIO', 'FECHA_REGISTRO', 
                                    'HORA_REGISTRO', 'DN_A_PORTAR', 'TIPO_PORTABILIDAD', 
                                    'RANGO_HORARIO']].to_excel(writer, sheet_name='Registros_Completos', index=False)
                
                output.seek(0)
                st.download_button(
                    label="📥 Descargar Reporte Excel",
                    data=output,
                    file_name=f'reporte_portabilidades_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
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
                    file_name=f'reporte_portabilidades_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv',
                    mime='text/csv'
                )
            except Exception as e:
                st.error(f"Error al generar CSV: {str(e)}")
        
        st.markdown("---")
        
        # Vista previa de datos a exportar
        st.write("**Vista previa del reporte a exportar:**")
        st.dataframe(reporte, use_container_width=True)
    else:
        st.info("No hay datos para exportar")

except Exception as e:
st.error(f"❌ Error general: {str(e)}")
st.error(traceback.format_exc())

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
