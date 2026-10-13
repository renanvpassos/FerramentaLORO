import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Processador de Planilhas", page_icon="📊", layout="centered")

st.title("📊 Processador de Múltiplas Planilhas para Importação")
st.markdown("Faça o upload de uma ou mais planilhas de origem.")

uploaded_files = st.file_uploader("Envie as planilhas de origem (Excel ou CSV)", type=["xlsx", "xls", "csv"], accept_multiple_files=True)

if uploaded_files:
    try:
        all_dfs = []
        
        for uploaded_file in uploaded_files:
            file_extension = uploaded_file.name.split('.')[-1].lower()
            
            # Carrega o arquivo sem cabeçalho fixo para permitirmos a identificação dinâmica de múltiplos headers
            if file_extension in ['xlsx', 'xls']:
                df_raw = pd.read_excel(uploaded_file, header=None)
            elif file_extension == 'csv':
                df_raw = pd.read_csv(uploaded_file, header=None)
            else:
                st.warning(f"Formato não suportado para o arquivo: {uploaded_file.name}")
                continue

            # Função auxiliar para verificar valores em uma linha
            def row_contains_terms(row, terms):
                row_str_list = [str(val).strip().lower() for val in row.values if pd.notna(val)]
                matches = sum(1 for term in terms if any(term in s for s in row_str_list))
                return matches >= 2  # Considera cabeçalho se encontrar pelo menos 2 termos-chave

            # Termos essenciais para identificar uma linha de cabeçalho
            search_terms = ["proforma", "material", "amount in brl"]

            # Identificar todas as linhas que atuam como cabeçalho
            header_indices = []
            for idx, row in df_raw.iterrows():
                if row_contains_terms(row, search_terms):
                    header_indices.append(idx)

            # Se nenhum cabeçalho foi encontrado, tenta usar o comportamento padrão (ex: linha 13)
            if not header_indices:
                header_indices = [13] if len(df_raw) > 13 else [0]

            # Encontrar a linha de parada ("GOODS SHIPPED") global no arquivo
            stop_index = None
            for idx, row in df_raw.iterrows():
                for val in row.values:
                    if val and "GOODS SHIPPED" in str(val).upper():
                        stop_index = idx
                        break
                if stop_index is not None:
                    break

            # Processar blocos entre os cabeçalhos encontrados
            dfs_arquivo = []
            for i, h_idx in enumerate(header_indices):
                # Onde o bloco começa (linha seguinte ao cabeçalho)
                start_row = h_idx + 1
                
                # Onde o bloco termina (próximo cabeçalho ou o limite de 'GOODS SHIPPED')
                if i + 1 < len(header_indices):
                    end_row = header_indices[i + 1]
                else:
                    end_row = stop_index if stop_index is not None else len(df_raw)

                if start_row >= end_row:
                    continue

                # Extrai o sub-dataframe usando a linha h_idx como cabeçalho
                df_bloco = df_raw.iloc[start_row:end_row].copy()
                df_bloco.columns = df_raw.iloc[h_idx].values

                def find_column(df, target):
                    for col in df.columns:
                        if col and str(col).strip().lower() == target.strip().lower():
                            return col
                    return None

                col_proforma = find_column(df_bloco, "Proforma")
                col_material = find_column(df_bloco, "Material")
                col_amount = find_column(df_bloco, "amount in BRL")
                col_weight = find_column(df_bloco, "net weight in KG")
                col_qty = find_column(df_bloco, "Packed Quantity")

                if not (col_proforma and col_material and col_amount and col_weight and col_qty):
                    continue  # Pula blocos que não contenham todas as colunas necessárias

                df_novo = pd.DataFrame()
                df_novo["invoice"] = df_bloco[col_proforma]
                df_novo["partnumber"] = df_bloco[col_material]
                df_novo["precototal"] = df_bloco[col_amount]
                df_novo["pesototal"] = df_bloco[col_weight]
                df_novo["quantidade"] = df_bloco[col_qty]
                df_novo["MOEDA"] = 790
                df_novo["INCOTERMS"] = "FCA"

                # Remove linhas totalmente vazias ou que repetem o texto do cabeçalho
                df_novo = df_novo.dropna(subset=["invoice", "partnumber"], how="all")
                df_novo = df_novo[~df_novo["partnumber"].astype(str).str.lower().str.contains("material")]

                if not df_novo.empty:
                    dfs_arquivo.append(df_novo)

            if dfs_arquivo:
                df_file_final = pd.concat(dfs_arquivo, ignore_index=True)
                all_dfs.append(df_file_final)
            else:
                st.warning(f"O arquivo '{uploaded_file.name}' não pôde ser mapeado corretamente com os cabeçalhos encontrados.")

        if all_dfs:
            # Concatena todos os arquivos processados em sequência
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
