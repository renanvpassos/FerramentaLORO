import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Processador de Planilhas", page_icon="📊", layout="centered")

st.title("📊 Processador de Planilhas para Importação")
st.markdown("Faça o upload da planilha de origem para gerar o arquivo formatado conforme as regras especificadas.")

uploaded_file = st.file_uploader("Envie a planilha de origem (Excel ou CSV)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        # Lê o arquivo dependendo da extensão (cabeçalho começa na linha 14 -> header=13)
        file_extension = uploaded_file.name.split('.')[-1].lower()
        if file_extension in ['xlsx', 'xls']:
            df_origem = pd.read_excel(uploaded_file, header=13)
        elif file_extension == 'csv':
            df_origem = pd.read_csv(uploaded_file, header=13)
        else:
            st.error("Formato de arquivo não suportado.")
            st.stop()

        st.subheader("Pré-visualização dos dados brutos (Início na linha 14)")
        st.dataframe(df_origem.head())

        # Botão para processar
        if st.button("Processar Planilha"):
            df_novo = pd.DataFrame()
            
            def find_column(df, target):
                for col in df.columns:
                    if str(col).strip().lower() == target.strip().lower():
                        return col
                return None

            col_proforma = find_column(df_origem, "Proforma")
            col_material = find_column(df_origem, "Material")
            col_amount = find_column(df_origem, "amount in BRL")
            col_weight = find_column(df_origem, "net weight in KG")
            col_qty = find_column(df_origem, "Packed Quantity")

            missing_cols = []
            if not col_proforma: missing_cols.append("Proforma")
            if not col_material: missing_cols.append("Material")
            if not col_amount: missing_cols.append("amount in BRL")
            if not col_weight: missing_cols.append("net weight in KG")
            if not col_qty: missing_cols.append("Packed Quantity")

            if missing_cols:
                st.error(f"As seguintes colunas obrigatórias não foram encontradas na planilha: {', '.join(missing_cols)}")
                st.info(f"Colunas disponíveis detetadas: {list(df_origem.columns)}")
            else:
                # Filtrar para ignorar linhas a partir de onde aparece "GOODS SHIPPED INTO" em qualquer coluna
                stop_index = None
                for idx, row in df_origem.iterrows():
                    for val in row.values:
                        if val and "GOODS SHIPPED INTO" in str(val).upper():
                            stop_index = idx
                            break
                    if stop_index is not None:
                        break
                
                if stop_index is not None:
                    df_origem_filtrado = df_origem.iloc[:stop_index].copy()
                    st.info(f"Linhas abaixo de 'GOODS SHIPPED INTO' foram ignoradas (total de linhas válidas: {len(df_origem_filtrado)}).")
                else:
                    df_origem_filtrado = df_origem.copy()

                # Preenchendo as colunas do novo formato solicitado
                df_novo["invoice"] = df_origem_filtrado[col_proforma]
                df_novo["partnumber"] = df_origem_filtrado[col_material]
                df_novo["precototal"] = df_origem_filtrado[col_amount]
                df_novo["pesototal"] = df_origem_filtrado[col_weight]
                df_novo["quantidade"] = df_origem_filtrado[col_qty]
                df_novo["MOEDA"] = 790
                df_novo["INCOTERMS"] = "FCA"

                # Remove linhas vazias
                df_novo = df_novo.dropna(subset=["invoice", "partnumber"], how="all")

                st.success("Planilha processada com sucesso!")
                st.subheader("Pré-visualização do Resultado Formatado")
                st.dataframe(df_novo.head(10))

                # Opção de exportação para Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_novo.to_excel(writer, index=False, sheet_name='Tratado')
                processed_data = output.getvalue()

                st.download_button(
                    label="📥 Baixar Planilha Tratada (.xlsx)",
                    data=processed_data,
                    file_name="planilha_tratada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Opção de exportação para CSV
                csv_data = df_novo.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Baixar Planilha Tratada (.csv)",
                    data=csv_data,
                    file_name="planilha_tratada.csv",
                    mime="text/csv"
                )

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
