import os
import re
import random
from datetime import datetime, timedelta
import pdfplumber
from docx import Document

# Tenta importar num2words; caso não esteja instalado, utiliza função de fallback
try:
    from num2words import num2words
    HAS_NUM2WORDS = True
except ImportError:
    HAS_NUM2WORDS = False

# Mapeamento de meses em português
MESES = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL",
    5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
    9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO"
}

# Constantes para conversão manual de números por extenso (fallback)
UNIDADES = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove",
            "dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"]
DEZENAS = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"]
CENTENAS = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"]

def _converter_grupo(n):
    if n == 0:
        return ""
    if n == 100:
        return "cem"
    c = n // 100
    d = (n % 100) // 10
    u = n % 10
    partes = []
    if c > 0:
        partes.append(CENTENAS[c])
    if d == 1:
        partes.append(UNIDADES[d * 10 + u])
    else:
        if d > 1:
            partes.append(DEZENAS[d])
        if u > 0:
            partes.append(UNIDADES[u])
    return " e ".join(partes)

def valor_para_extenso_manual(valor):
    """Converte valor monetário numérico para extenso em maiúsculas (Python Puro)."""
    valor = round(valor, 2)
    inteiro = int(valor)
    centavos = int(round((valor - inteiro) * 100))
    
    if valor == 0:
        return "ZERO REAIS"
    
    partes_int = []
    milhoes = inteiro // 1_000_000
    milhares = (inteiro % 1_000_000) // 1_000
    unidades = inteiro % 1_000

    if milhoes > 0:
        ext_m = _converter_grupo(milhoes)
        partes_int.append(f"{ext_m} {'milhão' if milhoes == 1 else 'milhões'}")
    if milhares > 0:
        ext_k = _converter_grupo(milhares)
        if ext_k == "um":
            partes_int.append("um mil")
        else:
            partes_int.append(f"{ext_k} mil")
    if unidades > 0:
        ext_u = _converter_grupo(unidades)
        partes_int.append(ext_u)

    txt_int = ""
    if partes_int:
        if len(partes_int) > 1:
            txt_int = " ".join(partes_int[:-1]) + " e " + partes_int[-1]
        else:
            txt_int = partes_int[0]
            
        txt_int += " real" if inteiro == 1 else " reais"

    txt_cent = ""
    if centavos > 0:
        ext_c = _converter_grupo(centavos)
        txt_cent = f"{ext_c} centavo" if centavos == 1 else f"{ext_c} centavos"

    if txt_int and txt_cent:
        resultado = f"{txt_int} e {txt_cent}"
    elif txt_int:
        resultado = txt_int
    else:
        resultado = txt_cent

    return resultado.upper()

def valor_por_extenso(valor):
    """Retorna o valor por extenso em MAIÚSCULAS usando num2words ou fallback."""
    if HAS_NUM2WORDS:
        try:
            return num2words(valor, lang='pt_BR', to='currency').upper()
        except Exception:
            pass
    return valor_para_extenso_manual(valor)

def subtrair_dias_uteis(data, dias=2):
    """Subtrai dias úteis ignorando sábados e domingos."""
    dias_subtraidos = 0
    data_atual = data
    while dias_subtraidos < dias:
        data_atual -= timedelta(days=1)
        if data_atual.weekday() < 5:  # 0-4 é Segunda a Sexta
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
            texto_completo += (page.extract_text() or "") + "\n"

            tables = page.extract_tables()
            for table in tables:
                for row in table:
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

def reajustar_valores(itens, percentual=0.03):
    """
    Aumenta o valor total no percentual informado, distribuindo variações
    aleatórias nos valores unitários dos itens.
    """
    total_original = sum(item["vlr_total"] for item in itens)
    total_alvo = round(total_original * (1 + percentual), 2)
    
    novos_itens = []
    soma_provisoria = 0.0

    for item in itens:
        min_fator = 1.0 + max(0.005, percentual - 0.02)
        max_fator = 1.0 + percentual + 0.02
        fator_aleatorio = random.uniform(min_fator, max_fator)
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

    diferenca = round(total_alvo - soma_provisoria, 2)
    if novos_itens and diferenca != 0:
        ultimo = novos_itens[-1]
        ultimo["vlr_total"] = round(ultimo["vlr_total"] + diferenca, 2)
        if ultimo["qtd"] > 0:
            ultimo["vlr_unit"] = round(ultimo["vlr_total"] / ultimo["qtd"], 2)

    return novos_itens, total_alvo

def preencher_tabela(tabela, novos_itens):
    """Identifica dinamicamente a posição das colunas no cabeçalho e preenche a tabela."""
    if not tabela.rows:
        return

    col_idx = {"num": 0, "desc": 1, "und": 2, "qtd": 3, "unit": 4, "total": 5}
    header_cells = [c.text.upper() for c in tabela.rows[0].cells]
    
    for idx, text in enumerate(header_cells):
        if "ORDEM" in text or "ITEM" in text or "Nº" in text:
            col_idx["num"] = idx
        elif "UND" in text or "UNID" in text:
            col_idx["und"] = idx
        elif "DESC" in text or "PRODUTO" in text or "SERVIÇO" in text:
            col_idx["desc"] = idx
        elif "QTD" in text or "QUANT" in text:
            col_idx["qtd"] = idx
        elif "UNIT" in text or "PR. UNIT" in text:
            col_idx["unit"] = idx
        elif "TOTAL" in text or "PR. TOTAL" in text:
            col_idx["total"] = idx

    for idx, item in enumerate(novos_itens):
        row_idx = idx + 1
        if row_idx >= len(tabela.rows):
            tabela.add_row()
        
        row_cells = tabela.rows[row_idx].cells
        if len(row_cells) > max(col_idx.values()):
            row_cells[col_idx["num"]].text = str(idx + 1)
            row_cells[col_idx["desc"]].text = str(item["descricao"])
            row_cells[col_idx["und"]].text = str(item["und"])
            row_cells[col_idx["qtd"]].text = f"{item['qtd']:.0f}" if item['qtd'].is_integer() else f"{item['qtd']:.2f}".replace(".", ",")
            row_cells[col_idx["unit"]].text = f"{item['vlr_unit']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            row_cells[col_idx["total"]].text = f"{item['vlr_total']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def aplicar_fonte_documento(doc, nome_fonte="Mongolian Baiti"):
    """
    Garante que a fonte em todo o documento (parágrafos, estilo padrão e tabelas)
    seja alterada para Mongolian Baiti.
    """
    if 'Normal' in doc.styles:
        doc.styles['Normal'].font.name = nome_fonte

    for p in doc.paragraphs:
        for run in p.runs:
            run.font.name = nome_fonte

    for tabela in doc.tables:
        for row in tabela.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        run.font.name = nome_fonte

def preencher_word(template_path, output_path, dados, novos_itens, total_geral, data_formatada, e_template_jw=False):
    """Preenche o arquivo Word de acordo com os placeholders e aplica a formatação adequada."""
    doc = Document(template_path)
    val_str = f"{total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    extenso_str = valor_por_extenso(total_geral)

    # 1. Preenchimento de parágrafos
    for p in doc.paragraphs:
        p_upper = p.text.upper()

        if "FORTALEZA" in p_upper and ("X DE X DE X" in p_upper or "XX DE XX DE XXXX" in p_upper or "XX DE XX DE X" in p_upper):
            if p.text.startswith("FORTALEZA") or p.text.startswith("Fortaleza"):
                prefixo = "FORTALEZA" if p.text.isupper() else "Fortaleza"
                p.text = f"{prefixo}, {data_formatada}"

        if "CLIENTE:" in p_upper and dados.get("cliente"):
            p.text = f"Cliente: {dados['cliente']}"
        if "CNPJ:" in p_upper and dados.get("cnpj") and not e_template_jw:
            p.text = f"CNPJ: {dados['cnpj']}"

        if e_template_jw and "VALOR DA PROPOSTA" in p_upper:
            p.text = f"Valor da Proposta R$ {val_str} ({extenso_str})"
        elif not e_template_jw and "VALOR TOTAL DA PROPOSTA:" in p_upper:
            p.text = f"VALOR TOTAL DA PROPOSTA:  R$ {val_str}"

    # 2. Preenchimento de Tabelas
    if doc.tables:
        tabela = doc.tables[0]
        preencher_tabela(tabela, novos_itens)

        # Se for o template da JW, aplica NEGRITO em TODAS as células da tabela (cabeçalho e dados)
        if e_template_jw:
            for row in tabela.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for run in p.runs:
                            run.bold = True

    # 3. Aplicação da Fonte Mongolian Baiti em TUDO
    aplicar_fonte_documento(doc, "Mongolian Baiti")

    doc.save(output_path)

def processar_notas():
    pasta_entrada = "notas_entrada"
    pasta_saida = "propostas_saida"
    
    template_lc = "template_lc.docx" if os.path.exists("template_lc.docx") else "template.docx"
    template_jw = "template_jw.docx"

    if not os.path.exists(pasta_entrada):
        os.makedirs(pasta_entrada)
        print(f"Pasta '{pasta_entrada}' criada. Adicione os PDFs nela e execute novamente.")
        return

    if not os.path.exists(pasta_saida):
        os.makedirs(pasta_saida)

    if not os.path.exists(template_lc):
        print(f"Aviso: Template '{template_lc}' não encontrado!")
    if not os.path.exists(template_jw):
        print(f"Aviso: Template '{template_jw}' não encontrado!")

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

        # Pasta de saída da NF
        nome_pasta_nf = f"NF {num_nf}"
        caminho_pasta_nf = os.path.join(pasta_saida, nome_pasta_nf)
        os.makedirs(caminho_pasta_nf, exist_ok=True)

        # -------------------------------------------------------------
        # 3. PROPOSTA 1: LC COMERCIAL (Vencedora +3%)
        # -------------------------------------------------------------
        if os.path.exists(template_lc):
            novos_itens_lc, total_lc = reajustar_valores(dados["itens"], percentual=0.03)
            caminho_lc_final = os.path.join(caminho_pasta_nf, f"LC COMERCIAL NF {num_nf}.docx")
            preencher_word(template_lc, caminho_lc_final, dados, novos_itens_lc, total_lc, data_formatada, e_template_jw=False)
            print(f"  ✓ Proposta LC criada: {caminho_lc_final} (+3.0%)")

        # -------------------------------------------------------------
        # 4. PROPOSTA 2: JW COMERCIAL (2ª Perdedora +5% a +8%)
        # -------------------------------------------------------------
        if os.path.exists(template_jw):
            margem_jw = random.uniform(0.05, 0.08)
            novos_itens_jw, total_jw = reajustar_valores(dados["itens"], percentual=margem_jw)
            caminho_jw_final = os.path.join(caminho_pasta_nf, f"JW COMERCIAL NF {num_nf}.docx")
            preencher_word(template_jw, caminho_jw_final, dados, novos_itens_jw, total_jw, data_formatada, e_template_jw=True)
            print(f"  ✓ Proposta JW criada: {caminho_jw_final} (+{margem_jw*100:.2f}%)")

if __name__ == "__main__":
    processar_notas()