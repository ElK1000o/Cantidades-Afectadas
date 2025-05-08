import streamlit as st
import pandas as pd
import re
import io

st.set_page_config(page_title="Resumen de Afectación", layout="centered")
st.title("📦 Análisis de Afectación en Unidades")

# --- Función para convertir código UM a unidades ---
def convertir_a_unidades(cantidad, um):
    if pd.isna(cantidad) or pd.isna(um):
        return 0

    try:
        cantidad = float(cantidad)
    except:
        return 0

    um = str(um).strip().upper()

    # 1Q = 1000 unidades
    match_q = re.match(r'^(\d*)Q$', um)
    if match_q:
        mult = int(match_q.group(1)) if match_q.group(1) else 1
        return cantidad * mult * 1000

    # 6 UN, 3 AJ
    match_unidad = re.match(r'^(\d*)\s*([A-Z]+)$', um)
    if match_unidad:
        mult, unidad = match_unidad.groups()
        mult = int(mult) if mult else 1
        if unidad == 'UN':
            return cantidad * mult
        elif unidad == 'AJ':
            return cantidad * mult  # ajustable si cada saco tiene 25, 50, etc.

    # Códigos tipo Y40, T00, etc.
    match_letra_num = re.match(r'^([A-Z])(\d{2})$', um)
    if match_letra_num:
        letra, numero = match_letra_num.groups()
        base = (ord('Z') - ord(letra) + 1) * 100
        return cantidad * (base + int(numero))

    return cantidad  # fallback

# --- Subida del archivo ---
archivo = st.file_uploader("📂 Sube tu archivo Excel con columnas: Código, Producto, Cantidad, UM, Cantidad Afectada, UM Afectada", type=["xlsx"])

if archivo is not None:
    try:
        df = pd.read_excel(archivo)

        columnas_necesarias = {'Código Producto', 'Nombre producto', 'Cantidad almacenada en bodega', 'Unidad de medida', 'Cantidad afectada cliente', 'Unidad de medida2'}
        if not columnas_necesarias.issubset(df.columns):
            st.error(f"Faltan columnas obligatorias: {columnas_necesarias}")
        else:
            # Convertir a unidades
            df['Cantidad_UM_Unidades'] = df.apply(lambda row: convertir_a_unidades(row['Cantidad almacenada en bodega'], row['Unidad de medida']), axis=1)
            df['Cantidad_Afectada_UM_Unidades'] = df.apply(lambda row: convertir_a_unidades(row['Cantidad afectada cliente'], row['Unidad de medida2']), axis=1)

            # Agrupar
            resumen = df.groupby(['Código Producto']).agg({
                'Cantidad_UM_Unidades': 'sum',
                'Cantidad_Afectada_UM_Unidades': 'sum'
            }).reset_index()

            resumen = resumen.rename(columns={
                'Cantidad_UM_Unidades': 'Cantidad en Unidades',
                'Cantidad_Afectada_UM_Unidades': 'Cantidad Afectada en Unidades'
            })

            resumen['% Afectado'] = round((resumen['Cantidad Afectada en Unidades'] / resumen['Cantidad en Unidades']) * 100, 2)

            # Mostrar resultados
            st.subheader("📊 Resumen final")
            st.dataframe(resumen, use_container_width=True)

            # Botón de descarga
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                resumen.to_excel(writer, index=False, sheet_name='Resumen')
                df.to_excel(writer, index=False, sheet_name='Detalle con Unidades')
            output.seek(0)

            st.download_button(
                label="📥 Descargar Excel con resultados",
                data=output,
                file_name="afectacion_convertida.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Ocurrió un error al procesar el archivo: {e}")
