import streamlit as st
import pandas as pd
import io

# Configuração da página web
st.set_page_config(page_title="Tratador de Dados de Municípios", layout="centered")

st.title("🏙️ Processador de Serviços Municipais")
st.write("Insira os dados abaixo para gerar sua planilha do Excel.")

# 1. Abre caixa na tela para digitar o nome do município
nome_municipio = st.text_input("Digite o nome do município:", placeholder="Ex: São Paulo")

# 2. Caixa de texto grande na própria página web
dados_brutos = st.text_area("Cole os dados a serem tratados abaixo:", height=300, placeholder="Cole as linhas aqui...")

# Botão para processar
if st.button("Processar Dados e Criar Planilha"):
    
    if not nome_municipio:
        st.warning("⚠️ Por favor, informe o nome do município!")
    elif not dados_brutos.strip():
        st.warning("⚠️ Nenhum dado foi colado para processamento.")
    else:
        # 3. Processamento das linhas (Igual ao seu código original)
        linhas = dados_brutos.strip().split("\n")
        dados_limpos = []
        item_atual = None
        descricao_atual = []
        aliquota_atual = None

        for linha in linhas:
            if "===" in linha or "---" in linha or "ITEM" in linha:
                continue

            # Divide por barras verticais e limpa espaços
            colunas = [c.strip() for c in linha.split("|")][1:-1]
            if not colunas or len(colunas) < 2:
                continue

            item = colunas[0]
            descricao = colunas[1]
            aliquota = colunas[2] if len(colunas) > 2 else ""

            if item:
                if item_atual:
                    dados_limpos.append([item_atual, " ".join(descricao_atual), aliquota_atual])
                item_atual = item
                descricao_atual = [descricao]
                aliquota_atual = aliquota
            else:
                if descricao:
                    descricao_atual.append(descricao)

        if item_atual:
            dados_limpos.append([item_atual, " ".join(descricao_atual), aliquota_atual])

        # Se encontrou dados válidos
        if dados_limpos:
            # 4. Criação do DataFrame
            df = pd.DataFrame(dados_limpos, columns=["ITEM", "DESCRIÇÃO DOS SERVIÇOS", "ALÍQUOTA"])

            # Mostra uma prévia dos dados na tela da web
            st.success("🎉 Dados processados com sucesso!")
            st.subheader("Prévia dos Dados:")
            st.dataframe(df.head(10)) # Mostra as 10 primeiras linhas

            # 5. Salvar na planilha na memória para o usuário baixar pelo navegador
            nome_arquivo = f"servicos_{nome_municipio.lower().replace(' ', '_')}.xlsx"
            nome_guia = nome_municipio[:31]

            # Cria o arquivo do Excel na memória (necessário para ambiente web)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=nome_guia, index=False)
            
            # Botão para o usuário fazer o download do arquivo gerado
            st.download_button(
                label="📥 Baixar Planilha do Excel",
                data=buffer.getvalue(),
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ Não foi possível extrair nenhum dado válido do texto colado. Verifique o formato.")
