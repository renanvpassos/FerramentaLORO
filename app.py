import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Processador de Planilhas", page_icon="📊", layout="centered")

st.title("📊 Processador de Múltiplas Planilhas para Importação")
st.markdown("Faça o upload de uma ou mais planilhas de origem. Os dados serão consolidados em uma única tabela em sequência, sem linhas em branco.")

uploaded_files = st.file_uploader("Envie as planilhas de origem (Excel ou CSV)", type=["xlsx", "xls", "csv"], accept_multiple_files=True)

if uploaded_files:
    try:
        all_dfs = []
        
        for uploaded_file in uploaded_files:
            file_extension = uploaded_file.name.split('.')[-1].lower()
            if file_extension in ['xlsx', 'xls']:
                df_origem = pd.read_excel(uploaded_file, header=13)
            elif file_extension == 'csv':
                df_origem = pd.read_csv(uploaded_file, header=13)
            else:
                st.warning(f"Formato não suportado para o arquivo: {uploaded_file.name}")
                continue

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
                st.error(f"O arquivo '{uploaded_file.name}' não possui as colunas: {', '.join(missing_cols)}")
                st.stop()

            # Filtrar para ignorar linhas a partir de onde aparece "GOODS SHIPPED INTO"
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
            else:
                df_origem_filtrado = df_origem.copy()

            # Cria o DataFrame estruturado para este arquivo
            df_novo = pd.DataFrame()
            df_novo["invoice"] = df_origem_filtrado[col_proforma]
            df_novo["partnumber"] = df_origem_filtrado[col_material]
            df_novo["precototal"] = df_origem_filtrado[col_amount]
            df_novo["pesototal"] = df_origem_filtrado[col_weight]
            df_novo["quantidade"] = df_origem_filtrado[col_qty]
            df_novo["MOEDA"] = 790
            df_novo["INCOTERMS"] = "FCA"

            # Remove linhas totalmente vazias para evitar espaçamento
            df_novo = df_novo.dropna(subset=["invoice", "partnumber"], how="all")
            
            if not df_novo.empty:
                all_dfs.append(df_novo)

        if all_dfs:
            # Concatena todos os dataframes em sequência, sem pular linhas
            df_final = pd.concat(all_dfs, ignore_index=True)

            st.success(f"Foram processados {len(uploaded_files)} arquivo(s) com sucesso! Total de linhas consolidadas: {len(df_final)}")
            st.subheader("Pré-visualização do Resultado Consolidado")
            st.dataframe(df_final.head(15))

            # Opção de exportação para Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Consolidado')
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Baixar Planilha Consolidada (.xlsx)",
                data=processed_data,
                file_name="planilha_consolidada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Opção de exportação para CSV
            csv_data = df_final.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Baixar Planilha Consolidada (.csv)",
                data=csv_data,
                file_name="planilha_consolidada.csv",
                mime="text/csv"
            )
        else:
            st.warning("Nenhum dado válido encontrado para consolidar nos arquivos enviados.")

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar os arquivos: {e}")
