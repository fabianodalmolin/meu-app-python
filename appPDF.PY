import streamlit as st
import pandas as pd
import pdfplumber
import io

# Configuração da página web
st.set_page_config(page_title="Processador de PDF de Alíquotas", layout="centered")

st.title("📄 Processador de Serviços Municipais (via PDF)")
st.write("Faça o upload do arquivo PDF contendo a tabela de serviços para gerar sua planilha do Excel.")

# 1. Abre caixa na tela para digitar o nome do município
nome_municipio = st.text_input("Digite o nome do município:", placeholder="Ex: Florianópolis")

# 2. Componente de upload para receber o arquivo PDF do usuário
arquivo_pdf = st.file_uploader("Selecione o arquivo PDF da lista de serviços:", type=["pdf"])

# Botão para processar
if st.button("Processador Dados do PDF e Criar Planilha"):
    
    if not nome_municipio:
        st.warning("⚠️ Por favor, informe o nome do município!")
    elif arquivo_pdf is None:
        st.warning("⚠️ Por favor, faça o upload de um arquivo PDF válido.")
    else:
        dados_limpos = []
        
        # 3. Extração e leitura de texto de todas as páginas do PDF enviado
        with pdfplumber.open(arquivo_pdf) as pdf:
            for pagina in pdf.pages:
                texto_pagina = pagina.extract_text()
                if not texto_pagina:
                    continue
                
                linhas = texto_pagina.split("\n")
                
                item_atual = None
                descricao_atual = []
                aliquota_atual = None

                for linha in linhas:
                    # Limpa espaços nas pontas da linha
                    linha_limpa = linha.strip()
                    
                    # Pula cabeçalhos da página ou linhas divisórias
                    if not linha_limpa or "Subitem" in linha_limpa or "Descrição" in linha_limpa or "Alíquota" in list(linha_limpa):
                        continue
                    
                    # Divide a linha tentando identificar o código numérico no início (ex: 1.01, 11.04)
                    partes = linha_limpa.split(" ", 1)
                    primeira_palavra = partes[0]
                    
                    # Verifica se a linha começa com um número de subitem válido (ex: 1.01 ou 1) ou se é continuação
                    # Checa se possui pontos ou se é um número inteiro de grupo principal (ex: "1", "2")
                    eh_codigo = (primeira_palavra.replace(".", "").isdigit() and len(partes) > 1)
                    
                    if eh_codigo:
                        # Se já havia um item sendo processado, salva ele antes de começar o novo
                        if item_atual:
                            dados_limpos.append([item_atual, " ".join(descricao_atual), aliquota_atual])
                        
                        item_atual = primeira_palavra
                        resto_linha = partes[1].strip()
                        
                        # Tenta capturar a alíquota no final da linha (ex: "2,0%" ou "3,5%")
                        if resto_linha.endswith("%") or resto_linha.endswith("-"):
                            partes_fim = resto_linha.rsplit(" ", 1)
                            descricao_atual = [partes_fim[0].strip()]
                            aliquota_atual = partes_fim[1].strip()
                        else:
                            descricao_atual = [resto_linha]
                            aliquota_atual = ""
                    else:
                        # Se a linha não começa com número, ela é a continuação da descrição do item anterior
                        # Verifica se o final da linha trouxe a alíquota pendente do item
                        if linha_limpa.endswith("%") or linha_limpa.endswith("-"):
                            partes_fim = linha_limpa.rsplit(" ", 1)
                            if len(partes_fim) > 1:
                                descricao_atual.append(partes_fim[0].strip())
                                aliquota_atual = partes_fim[1].strip()
                            else:
                                aliquota_atual = linha_limpa
                        else:
                            descricao_atual.append(linha_limpa)

                # Salva o último item processado da página
                if item_atual:
                    dados_limpos.append([item_atual, " ".join(descricao_atual), aliquota_atual])

        # Se encontrou dados válidos extraídos do PDF
        if dados_limpos:
            # 4. Criação do DataFrame estruturado
            df = pd.DataFrame(dados_limpos, columns=["ITEM", "DESCRIÇÃO DOS SERVIÇOS", "ALÍQUOTA"])
            
            # Limpeza fina: remove linhas duplicadas ou vazias geradas no processo
            df = df.drop_duplicates().dropna(subset=["ITEM"])
            df = df[df["ITEM"] != ""]

            st.success("🎉 PDF processado com sucesso!")
            st.subheader("Prévia das Alíquotas Encontradas:")
            st.dataframe(df.head(15)) # Exibe as 15 primeiras linhas na tela

            # 5. Configuração da planilha em memória
            nome_arquivo = f"servicos_{nome_municipio.lower().replace(' ', '_')}.xlsx"
            nome_guia = nome_municipio[:31]

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=nome_guia, index=False)
            
            # Botão para o download imediato
            st.download_button(
                label="📥 Baixar Planilha do Excel",
                data=buffer.getvalue(),
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("❌ Não foi possível extrair dados estruturados deste PDF. Verifique se ele possui texto selecionável.")
