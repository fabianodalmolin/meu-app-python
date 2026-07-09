import streamlit as st
import pandas as pd
import pdfplumber
import io

# Configuração da página web
st.set_page_config(page_title="Processador de PDF de Alíquotas", layout="centered")

st.title("📄 Processador de Serviços Municipais (via PDF)")
st.write("Faça o upload do arquivo PDF contendo a tabela de serviços para gerar sua planilha do Excel.")

# 1. Abre caixa na tela para digitar o nome do município
nome_municipio = st.text_input("Digite o nome do município:", placeholder="Ex: Pomerode")

# 2. Componente de upload para receber o arquivo PDF do usuário
arquivo_pdf = st.file_uploader("Selecione o arquivo PDF da lista de serviços:", type=["pdf"])

# Botão para processar
if st.button("Processar Dados do PDF e Criar Planilha"):
    
    if not nome_municipio:
        st.warning("⚠️ Por favor, informe o nome do município!")
    elif arquivo_pdf is None:
        st.warning("⚠️ Por favor, faça o upload de um arquivo PDF válido.")
    else:
        dados_limpos = []
        
        # 3. Extração baseada nas linhas da tabela do PDF
        with pdfplumber.open(arquivo_pdf) as pdf:
            for pagina in pdf.pages:
                # Extrai os dados em formato de tabela (linhas e colunas estruturadas)
                tabela = pagina.extract_table()
                
                if not tabela:
                    continue
                
                for linha in tabela:
                    # Pula linhas totalmente vazias
                    if not linha:
                        continue
                    
                    # Converte todos os elementos da linha para texto limpo e remove nulos
                    colunas = [str(c).strip().replace('\n', ' ') if c is not None else "" for c in linha]
                    
                    # Filtra para garantir que a linha possui dados úteis
                    if len(colunas) >= 2:
                        item = colunas[0]
                        descricao = colunas[1]
                        # Pega a alíquota se houver uma terceira coluna, senão deixa em branco
                        aliquota = colunas[2] if len(colunas) > 2 else ""
                        
                        # Ignora os cabeçalhos das tabelas que se repetem nas páginas
                        if "Subitem" in item or "Descrição" in item or (item == "" and descricao == ""):
                            continue
                            
                        dados_limpos.append([item, descricao, aliquota])

        # Se encontrou dados válidos extraídos do PDF
        if dados_limpos:
            # 4. Criação do DataFrame estruturado (agora apenas com textos simples, sem listas!)
            df = pd.DataFrame(dados_limpos, columns=["ITEM", "DESCRIÇÃO DOS SERVIÇOS", "ALÍQUOTA"])
            
            # Remove possíveis linhas duplicadas com segurança
            df = df.drop_duplicates()

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
            st.error("❌ Não foi possível extrair dados estruturados deste PDF. Verifique se ele possui tabelas válidas.")
