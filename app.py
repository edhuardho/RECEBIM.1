import os
import subprocess
import sys

# Força a instalação das ferramentas direto pelo código se elas não existirem
try:
    import openpyxl
    import pandas as pd
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pandas", "openpyxl"])
    import openpyxl
    import pandas as pd

import streamlit as st
from datetime import datetime
import io

st.set_page_config(page_title="Controle de Recebimento", layout="centered")

st.title("📦 Sistema de Recebimento & Shelf Life")
st.write("Insira o CSV da Delicari, faça a conferência e gere a planilha oficial.")

if "dados_conferidos" not in st.session_state:
    st.session_state.dados_conferidos = []
if "csv_tratado" not in st.session_state:
    st.session_state.csv_tratado = None

st.subheader("📁 1. Upload de Arquivos")
arquivo_csv = st.file_uploader("Suba o arquivo CSV bruto (DELICARI)", type=["csv"], key="csv")
arquivo_modelo = st.file_uploader("Suba o seu modelo padrão (SHELF - PADRÃO.xlsx)", type=["xlsx"], key="modelo")

if arquivo_csv is not None and st.session_state.csv_tratado is None:
    try:
        df_bruto = pd.read_csv(arquivo_csv, sep="\t", encoding="utf-8", header=None)
        linhas_produtos = []
        
        for idx, row in df_bruto.iterrows():
            texto_linha = str(row.iloc[0])
            if " - " in texto_linha and ("IOGURTE" in texto_linha or "OGURTE" in texto_linha):
                texto_limpo = texto_linha.replace('"', '').replace('-', '').strip()
                partes = texto_limpo.split(' ', 1)
                if len(partes) == 2:
                    codigo = partes[0].strip().lstrip('0')
                    descricao = partes[1].replace('- ', '').strip()
                    linhas_produtos.append({"Codigo_Prod": codigo, "Descricao_Prod": descricao})
        
        st.session_state.csv_tratado = pd.DataFrame(linhas_produtos)
        st.success(f"✅ CSV processado! {len(linhas_produtos)} produtos encontrados.")
    except Exception as e:
        st.error(f"Erro ao processar CSV: {e}")

if st.session_state.csv_tratado is not None and arquivo_modelo is not None:
    st.markdown("---")
    st.subheader("📝 2. Conferência de Itens na Doca")
    
    df_atual = st.session_state.csv_tratado
    codigo_digitado = st.text_input("Digite ou Bipe o Código do Produto:").strip().lstrip('0')

    if codigo_digitado:
        item_encontrado = df_atual[df_atual['Codigo_Prod'] == codigo_digitado]

        if not item_encontrado.empty:
            descricao = item_encontrado.iloc[0]['Descricao_Prod']
            st.info(f"📦 **Produto Encontrado:** {descricao}")

            with st.form(key="form_lotes", clear_on_submit=True):
                col1, col2 = st.columns(2)
                with col1:
                    data_fab = st.date_input("Data de Fabricação", value=datetime.today())
                with col2:
                    data_ven = st.date_input("Data de Vencimento", value=datetime.today())
                
                qtd_lote = st.number_input("Quantidade deste Lote:", min_value=1, step=1)
                btn_salvar = st.form_submit_button("➕ Salvar Lote")

                if btn_salvar:
                    st.session_state.dados_conferidos.append({
                        "CODIGO": codigo_digitado,
                        "DESCRIÇÃO DO PRODUTO": descricao,
                        "FABRICAÇÃO": data_fab.strftime('%d/%m/%Y'),
                        "VALIDADE": data_ven.strftime('%d/%m/%Y'),
                        "QUANTIDADE": qtd_lote
                    })
                    st.toast(f"Lote salvo para o item {codigo_digitado}!", icon="✔️")
        else:
            st.error("⚠️ Código não encontrado neste recebimento.")

    if st.session_state.dados_conferidos:
        st.markdown("---")
        st.subheader("📋 Itens Conferidos")
        df_conferidos = pd.DataFrame(st.session_state.dados_conferidos)
        st.dataframe(df_conferidos)

        if st.button("🏁 Finalizar Recebimento e Criar Planilha"):
            try:
                wb = openpyxl.load_workbook(arquivo_modelo)
                aba_nome = "SHELF" if "SHELF" in wb.sheetnames else wb.sheetnames[0]
                ws = wb[aba_nome]
                
                linha_inicio = 5
                
                for index, row in df_conferidos.iterrows():
                    linha_atual = gateway_inicio = linha_inicio + index
                    ws[f"A{linha_atual}"] = int(row["CODIGO"]) if row["CODIGO"].isdigit() else row["CODIGO"]
                    ws[f"B{linha_atual}"] = row["DESCRIÇÃO DO PRODUTO"]
                    ws[f"C{linha_atual}"] = row["FABRICAÇÃO"]
                    ws[f"D{linha_atual}"] = row["VALIDADE"]
                    ws[f"H{linha_atual}"] = row["QUANTIDADE"]
                    
                    ws[f"E{linha_atual}"] = f"=D{linha_atual}-TODAY()"
                    ws[f"F{linha_atual}"] = f"=D{linha_atual}-C{linha_atual}"
                    ws[f"G{linha_atual}"] = f"=E{linha_atual}/F{linha_atual}"
                
                buffer = io.BytesIO()
                wb.save(buffer)
                buffer.seek(0)
                
                st.success("🎉 Planilha oficial gerada com sucesso!")
                st.download_button(
                    label="📥 Baixar Planilha de Controle Pronta",
                    data=buffer,
                    file_name=f"CONTROLE_SHELF_{datetime.now().strftime('%d%m%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Erro ao gerar o arquivo final: {e}")
