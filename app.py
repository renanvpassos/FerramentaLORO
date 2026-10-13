import streamlit as st
import pandas as pd
import re
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
            
            if file_extension in ['xlsx', 'xls']:
                df_raw = pd.read_excel(uploaded_file, header=None)
            elif file_extension == 'csv':
                df_raw = pd.read_csv(uploaded_file, header=None)
            else:
                st.warning(f"Formato não suportado para o arquivo: {uploaded_file.name}")
                continue

            # Identificar todas as linhas onde a Coluna A (índice 0) contém "Proforma"
            header_indices = []
            for idx, row in df_raw.iterrows():
                val_col_a = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ""
                if "proforma" in val_col_a:
                    header_indices.append(idx)

            if not header_indices:
                st.warning(f"O arquivo '{uploaded_file.name}' não possui nenhum cabeçalho com 'Proforma' na Coluna A.")
                continue

            # Mapear linhas de parada globais no arquivo (LORO... e GOODS SHIPPED)
            stop_points = []
            for idx, row in df_raw.iterrows():
                for val in row.values:
                    if pd.notna(val):
                        val_str = str(val).strip().upper()
                        if val_str.startswith("LORO") or "GOODS SHIPPED" in val_str:
                            stop_points.append(idx)
                            break

            dfs_arquivo = []
            for i, h_idx in enumerate(header_indices):
                start_row = h_idx + 1
                end_row = len(df_raw)
                
                if i + 1 < len(header_indices):
                    end_row = header_indices[i + 1]
                
                for s_idx in stop_points:
                    if start_row <= s_idx < end_row:
                        end_row = s_idx
                        break

                if start_row >= end_row:
                    continue

                df_bloco = df_raw.iloc[start_row:end_row].copy()
                
                raw_headers = df_raw.iloc[h_idx].values
                clean_headers = [str(col).strip() if pd.notna(col) else f"Unnamed_{c_idx}" for c_idx, col in enumerate(raw_headers)]
                df_bloco.columns = clean_headers

                # Função robusta que ignora quebras de linha, espaços extras e diferenças de maiúsculas/minúsculas
                def find_column(df, target):
                    target_clean = re.sub(r'\s+', '', target).lower()
                    for col in df.columns:
                        col_clean = re.sub(r'\s+', '', str(col)).lower()
                        if target_clean in col_clean:
                            return col
                    return None

                col_proforma = find_column(df_bloco, "Proforma")
                col_material = find_column(df_bloco, "Material")
                col_amount = find_column(df_bloco, "amount in BRL")
                col_weight = find_column(df_bloco, "net weight in KG")
                col_qty = find_column(df_bloco, "Packed Quantity")

                if not (col_proforma and col_material and col_amount and col_weight and col_qty):
                    continue  # Pula blocos vazios ou inválidos sem travar

                df_novo = pd.DataFrame()
                df_novo["invoice"] = df_bloco[col_proforma]
                df_novo["partnumber"] = df_bloco[col_material]
                df_novo["precototal"] = df_bloco[col_amount]
                df_novo["pesototal"] = df_bloco[col_weight]
                df_novo["quantidade"] = df_bloco[col_qty]
                df_novo["MOEDA"] = 790
                df_novo["INCOTERMS"] = "FCA"

                df_novo = df_novo.dropna(subset=["invoice", "partnumber"], how="all")
                df_novo = df_novo[~df_novo["partnumber"].astype(str).str.lower().str.contains("material")]

                if not df_novo.empty:
                    dfs_arquivo.append(df_novo)

            if dfs_arquivo:
                df_file_final = pd.concat(dfs_arquivo, ignore_index=True)
                all_dfs.append(df_file_final)
            else:
                st.warning(f"O arquivo '{uploaded_file.name}' não gerou dados válidos entre os limites mapeados.")

        if all_dfs:
            df_final = pd.concat(all_dfs, ignore_index=True)

            st.success(f"Foram processados {len(uploaded_files)} arquivo(s) com sucesso! Total de linhas consolidadas: {len(df_final)}")
            st.subheader("Pré-visualização do Resultado Consolidado")
            st.dataframe(df_final.head(15))

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
