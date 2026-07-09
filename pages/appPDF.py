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
                    # Remove linhas vazias ou cabeçalhos da tabela
                    if not linha or any(item is None for item in linha):
                        continue
                    
                    # Limpa os espaços em branco de cada coluna detectada
                    colunas = [str(c).strip().replace('\n', ' ') for c in linha]
                    
                    # Garante que temos pelo menos as colunas de Código e Descrição
                    if len(colunas) >= 2:
                        item = colunas[0]
                        descricao = colunas[1]
                        # Pega a alíquota se ela existir, senão deixa em branco
                        aliquota = colunas[2] if len(colunas) > 2 else ""
                        
                        # Filtra linhas inúteis ou cabeçalhos textuais duplicados
                        if "Subitem" in item or "Descrição" in item or item == "":
                            continue
                            
                        dados_limpos.append([item, list(descricao), aliquota])

        # Se encontrou dados válidos extraídos do PDF
        if dados_limpos:
            # 4. Criação do DataFrame estruturado
            df = pd.DataFrame(dados_limpos, columns=["ITEM", "DESCRIÇÃO DOS SERVIÇOS", "ALÍQUOTA"])
            
            # Remove possíveis linhas duplicadas para deixar a planilha limpa
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
            st.error("❌ Não foi possível extrair dados estruturados deste PDF. Verifique se ele possui uma tabela com linhas visíveis.")
