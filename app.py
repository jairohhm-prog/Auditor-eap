import streamlit as st
import openpyxl
from openpyxl.styles import Font
import io

# 1. CONFIGURAÇÃO DA INTERFACE (A CARA DO PROGRAMA)
st.set_page_config(page_title="Auditor de EAP", layout="centered")
st.title("📊 Automatizador e Auditor de EAP")
st.write("Faça o upload da sua planilha Excel. O sistema fará a numeração automática e corrigirá a hierarquia dos níveis.")

# 2. BOTÃO DE UPLOAD PARA O USUÁRIO FINAL
arquivo_carregado = st.file_uploader("Selecione a planilha (.xlsx)", type=["xlsx"])

if arquivo_carregado is not None:
    st.info("Arquivo recebido! Processando a itemização...")
    
    # Carrega a planilha a partir da memória, sem precisar salvar no PC
    wb = openpyxl.load_workbook(arquivo_carregado)
    planilha = wb.active

    # --- INÍCIO DA NOSSA LÓGICA MATEMÁTICA EXATA ---
    coluna_item = None
    coluna_descricao = None
    linha_inicio_dados = None

    # Radar de colunas
    for linha in range(1, 21):
        for coluna in range(1, planilha.max_column + 1):
            valor_celula = str(planilha.cell(row=linha, column=coluna).value).strip().upper()
            if valor_celula == "ITEM":
                coluna_item = coluna
                linha_inicio_dados = linha + 1 
            elif valor_celula in ["DESCRIÇÃO", "DESCRICAO"]:
                coluna_descricao = coluna
        if coluna_item is not None and coluna_descricao is not None:
            break

    if coluna_item is None or coluna_descricao is None:
        st.error("❌ ERRO: Não localizei as colunas 'ITEM' e 'DESCRIÇÃO' no cabeçalho.")
    else:
        hierarquia_atual = [] 
        prefixo_base_atual = ""
        contador_servico = 1
        log_correcoes = [] 

        # Barra de progresso visual para o usuário
        barra_progresso = st.progress(0)
        total_linhas = planilha.max_row - linha_inicio_dados + 1

        for idx, linha in enumerate(range(linha_inicio_dados, planilha.max_row + 1)):
            celula_descricao = planilha.cell(row=linha, column=coluna_descricao).value
            celula_item = planilha.cell(row=linha, column=coluna_item).value

            if celula_descricao is None or str(celula_descricao).strip() == "":
                continue

            if celula_item is not None and str(celula_item).strip() != "":
                item_digitado = str(celula_item).strip()
                partes = item_digitado.split('.')
                profundidade = len(partes)
                
                if profundidade > len(hierarquia_atual):
                    while len(hierarquia_atual) < profundidade:
                        hierarquia_atual.append(1)
                elif profundidade == len(hierarquia_atual):
                    hierarquia_atual[-1] += 1
                else:
                    hierarquia_atual = hierarquia_atual[:profundidade]
                    hierarquia_atual[-1] += 1
                    
                item_correto = ".".join(str(x) for x in hierarquia_atual)
                
                if item_correto != item_digitado:
                    descricao_texto = str(celula_descricao).strip()
                    log_correcoes.append((linha, descricao_texto, item_digitado, item_correto))
                    
                planilha.cell(row=linha, column=coluna_item).value = item_correto
                prefixo_base_atual = item_correto
                contador_servico = 1 
                
            else:
                if prefixo_base_atual != "":
                    novo_item = f"{prefixo_base_atual}.{contador_servico}"
                    planilha.cell(row=linha, column=coluna_item).value = novo_item
                    contador_servico += 1
            
            # Atualiza a barra de progresso
            barra_progresso.progress((idx + 1) / total_linhas)

        # Geração da aba de LOG
        if "LOG" in wb.sheetnames:
            wb.remove(wb["LOG"])
        aba_log = wb.create_sheet(title="LOG")
        
        if len(log_correcoes) > 0:
            aba_log.append(["Linha", "Descrição do Item", "O que estava digitado", "Como foi corrigido"])
            for celula in aba_log[1]:
                celula.font = Font(bold=True)
            for erro in log_correcoes:
                aba_log.append([erro[0], erro[1], erro[2], erro[3]])
            st.warning(f"⚠️ Atenção: Foram corrigidas {len(log_correcoes)} inconsistências na numeração. Verifique a aba LOG no arquivo baixado.")
        else:
            aba_log.append(["Nenhuma correção de nível foi necessária na estrutura da EAP."])
            st.success("✅ Estrutura perfeita! Nenhuma correção foi necessária.")

        # --- PREPARAÇÃO DO ARQUIVO PARA DOWNLOAD ---
        # Salva o resultado em uma memória temporária para enviar ao usuário
        saida_memoria = io.BytesIO()
        wb.save(saida_memoria)
        saida_memoria.seek(0) # Volta ao início do arquivo na memória

        # 3. BOTÃO DE DOWNLOAD
        st.download_button(
            label="📥 Baixar Planilha Itemizada e Auditada",
            data=saida_memoria,
            file_name="EAP_Automatizada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )