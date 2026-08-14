import os
import re
import random
from datetime import datetime, timedelta
import pdfplumber
from docx import Document

# Mapeamento de meses em português
MESES = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL",
    5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
    9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"
}

def subtrair_dias_uteis(data, dias=2):
    """Subtrai dias úteis ignorando sábados e domingos."""
    dias_subtraidos = 0
    data_atual = data
    while dias_subtraidos < dias:
        data_atual -= timedelta(days=1)
        # 0 = Segunda, ..., 4 = Sexta, 5 = Sábado, 6 = Domingo
        if data_atual.weekday() < 5:
            dias_subtraidos += 1
    return data_atual

def formatar_data_word(data_obj):
    """Formata a data no padrão: DD DE MÊS DE YYYY"""
    dia = data_obj.strftime("%d")
    mes = MESES[data_obj.month]
    ano = data_obj.year
    return f"{dia} DE {mes} DE {ano}"

def extrair_dados_pdf(caminho_pdf):
    """Extrai informações relevantes e itens de todas as páginas da NF."""
    dados = {
        "numero_nf": None,
        "cliente": None,
        "cnpj": None,
        "data_emissao": None,
        "itens": []
    }

    with pdfplumber.open(caminho_pdf) as pdf:
        texto_completo = ""
        for page in pdf.pages:
            texto_completo += page.extract_text() + "\n"

            # Extração de tabelas (produtos)
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    # Filtra linhas válidas de produtos (geralmente começam com código numérico)
                    if row and len(row) >= 9 and row[0] and row[0].isdigit():
                        descricao = row[1].replace("\n", " ") if row[1] else ""
                        unidade = row[5] if row[5] else ""
                        
                        try:
                            qtd = float(row[6].replace(".", "").replace(",", "."))
                            vlr_unit = float(row[7].replace(".", "").replace(",", "."))
                            vlr_total = float(row[8].replace(".", "").replace(",", "."))
                            
                            dados["itens"].append({
                                "descricao": descricao,
                                "und": unidade,
                                "qtd": qtd,
                                "vlr_unit": vlr_unit,
                                "vlr_total": vlr_total
                            })
                        except (ValueError, TypeError):
                            continue

        # Extração de Cabeçalho via Regex
        nf_match = re.search(r"Nº\s*(\d+)", texto_completo)
        if nf_match:
            dados["numero_nf"] = nf_match.group(1)

        data_match = re.search(r"DATA DE EMISSÃO\n*(\d{2}/\d{2}/\d{4})", texto_completo)
        if not data_match:
            data_match = re.search(r"(\d{2}/\d{2}/\d{4})", texto_completo)
        if data_match:
            dados["data_emissao"] = datetime.strptime(data_match.group(1), "%d/%m/%Y")

        cliente_match = re.search(r"NOME/RAZÃO SOCIAL\n*(.*?)\n", texto_completo)
        if cliente_match:
            dados["cliente"] = cliente_match.group(1).strip()

        cnpj_match = re.search(r"CNPJ/CPF\n*(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", texto_completo)
        if cnpj_match:
            dados["cnpj"] = cnpj_match.group(1).strip()

    return dados

def reajustar_valores(itens):
    """
    Aumenta o valor total em exatamente 3%, distribuindo variações aleatórias
    nos valores unitários dos itens.
    """
    total_original = sum(item["vlr_total"] for item in itens)
    total_alvo = round(total_original * 1.03, 2)
    
    novos_itens = []
    soma_provisoria = 0.0

    for i, item in enumerate(itens):
        # Variação individual entre +1% e +5%
        fator_aleatorio = random.uniform(1.01, 1.05)
        novo_vlr_unit = round(item["vlr_unit"] * fator_aleatorio, 2)
        novo_vlr_total = round(novo_vlr_unit * item["qtd"], 2)

        novos_itens.append({
            "descricao": item["descricao"],
            "und": item["und"],
            "qtd": item["qtd"],
            "vlr_unit": novo_vlr_unit,
            "vlr_total": novo_vlr_total
        })
        soma_provisoria += novo_vlr_total

    # Ajuste de arredondamento no último item para fechar EXATAMENTE em +3%
    diferenca = round(total_alvo - soma_provisoria, 2)
    if novos_itens and diferenca != 0:
        ultimo = novos_itens[-1]
        ultimo["vlr_total"] = round(ultimo["vlr_total"] + diferenca, 2)
        ultimo["vlr_unit"] = round(ultimo["vlr_total"] / ultimo["qtd"], 2)

    return novos_itens, total_alvo

def preencher_word(template_path, output_path, dados, novos_itens, total_geral, data_formatada):
    doc = Document(template_path)

    # 1. Substituição de Placeholders nos parágrafos
    for p in doc.paragraphs:
        if "FORTALEZA, X DE X DE X" in p.text:
            p.text = f"FORTALEZA, {data_formatada}"
        if "Cliente:" in p.text and dados["cliente"]:
            p.text = f"Cliente: {dados['cliente']}"
        if "CNPJ:" in p.text and dados["cnpj"]:
            p.text = f"CNPJ: {dados['cnpj']}"
        if "VALOR TOTAL DA PROPOSTA:" in p.text:
            val_str = f"{total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            p.text = f"VALOR TOTAL DA PROPOSTA:  R$ {val_str}"

    # 2. Preenchimento da Tabela
    if doc.tables:
        tabela = doc.tables[0]
        
        # Identifica a linha de cabeçalho
        linhas_existentes = len(tabela.rows)
        # Primeira linha útil de dados é o índice 1 (logo após o cabeçalho)
        
        for idx, item in enumerate(novos_itens):
            row_idx = idx + 1 # +1 para pular o cabeçalho
            
            # Adiciona linha se faltar
            if row_idx >= len(tabela.rows):
                tabela.add_row()
            
            row_cells = tabela.rows[row_idx].cells
            
            # Preenche apenas nas colunas existentes (preservando o layout)
            if len(row_cells) >= 6:
                row_cells[0].text = str(idx + 1)
                row_cells[1].text = str(item["descricao"])
                row_cells[2].text = str(item["und"])
                row_cells[3].text = f"{item['qtd']:.0f}" if item['qtd'].is_integer() else f"{item['qtd']:.2f}".replace(".", ",")
                row_cells[4].text = f"{item['vlr_unit']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                row_cells[5].text = f"{item['vlr_total']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    doc.save(output_path)

def processar_notas():
    pasta_entrada = "notas_entrada"
    pasta_saida = "propostas_saida"
    template_file = "template.docx"

    if not os.path.exists(pasta_entrada):
        os.makedirs(pasta_entrada)
        print(f"Pasta '{pasta_entrada}' criada. Adicione os PDFs nela e execute novamente.")
        return

    if not os.path.exists(pasta_saida):
        os.makedirs(pasta_saida)

    arquivos = [f for f in os.listdir(pasta_entrada) if f.lower().endswith('.pdf')]
    
    if not arquivos:
        print("Nenhum arquivo PDF encontrado em 'notas_entrada'.")
        return

    for arquivo in arquivos:
        caminho_pdf = os.path.join(pasta_entrada, arquivo)
        print(f"\n--- Processando: {arquivo} ---")

        # 1. Leitura do PDF
        dados = extrair_dados_pdf(caminho_pdf)
        num_nf = dados["numero_nf"] or "000"

        # 2. Cálculo da Data (-2 dias úteis)
        data_emissao = dados["data_emissao"] or datetime.now()
        data_proposta = subtrair_dias_uteis(data_emissao, dias=2)
        data_formatada = formatar_data_word(data_proposta)

        # 3. Reajuste de Preços (+3%)
        novos_itens, total_geral = reajustar_valores(dados["itens"])

        # 4. Estruturação das Pastas de Saída
        nome_pasta_nf = f"NF {num_nf}"
        caminho_pasta_nf = os.path.join(pasta_saida, nome_pasta_nf)
        os.makedirs(caminho_pasta_nf, exist_ok=True)

        nome_arquivo_word = f"LC COMERCIAL NF {num_nf}.docx"
        caminho_word_final = os.path.join(caminho_pasta_nf, nome_arquivo_word)

        # 5. Gerar arquivo Word
        preencher_word(template_file, caminho_word_final, dados, novos_itens, total_geral, data_formatada)
        
        print(f"✓ Sucesso! Criado: {caminho_word_final}")

if __name__ == "__main__":
    processar_notas()